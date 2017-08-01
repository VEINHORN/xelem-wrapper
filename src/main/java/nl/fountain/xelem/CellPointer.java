/*
 * Created on 28-nov-2004
 * Copyright (C) 2004  Henk van den Berg
 *
 * This program is free software; you can redistribute it and/or
 * modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation; either version 2
 * of the License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program; if not, write to the Free Software
 * Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.
 * 
 * see license.txt
 *
 */
package nl.fountain.xelem;

import nl.fountain.xelem.excel.Worksheet;


/**
 * Keeps track of the position of cells being added to the
 * {@link nl.fountain.xelem.excel.Worksheet}.
 * <P>
 * [The position of the cellpointer is <em>not</em> displayed in the
 * actual Excel worksheet. If you wish to set the active cell in the
 * actual Excel worksheet on a position other then row 1, column 1, you
 * should use
 * {@link nl.fountain.xelem.excel.WorksheetOptions#setActiveCell(int, int)}.]
 * <P>
 * The position of the cellpointer is reflected in it's
 * {@link #getRowIndex() getRowIndex}- and 
 * {@link #getColumnIndex() getColumnIndex}-values.
 * The cellpointer moves to a new position relative to its old
 * position at a call to {@link #move}. In what direction it moves and over
 * how many cells depends on it's settings. The default setting is to step
 * 1 column to the right.
 */
public class CellPointer extends Address {
    
    /**
     * A constant for the method {@link #setMovement(int)}.
     */
    public static final int MOVE_RIGHT = 0;
    
    /**
     * A constant for the method {@link #setMovement(int)}.
     */
    public static final int MOVE_LEFT = 1;
    
    /**
     * A constant for the method {@link #setMovement(int)}.
     */
    public static final int MOVE_DOWN = 2;
    
    /**
     * A constant for the method {@link #setMovement(int)}.
     */
    public static final int MOVE_UP = 3;
    
    private int hStep;
    private int vStep;
    private int hMove;
    private int vMove;
    
    /**
     * Constructs a new CellPointer. The position of the pointer will be at
     * row 1, column 1. The initial movement of the pointer is MOVE_RIGHT.
     * The initial step distance of the pointer is 1. So the default behavior 
     * of this pointer at a call to {@link #move()} is to move one
     * column to the right.
     * 
     * @see nl.fountain.xelem.excel.Worksheet#getCellPointer()
     */
    public CellPointer() {
        super(1, 1);
        hStep = 1;
        vStep = 1;
        hMove = 1;
    }
    
    /**
     * Gets the address of the cell where this cellpointer is pointing at.
     * 
     * @return A new Address reflecting the momentary row- and column index
     * 			of this CellPointer.
     */
    public Address getAddress() {
        return new Address(r, c);
    }
    
    /**
     * Sets the number of cells this cellpointer will move 
     * in the horizontal axis after a call to {@link #move()}.
     * The default is 1.
     * <P>
     * If the movement of this cellpointer is set to
     * MOVE_DOWN or MOVE_UP the value of horizontalStepDistance
     * has no influence on this pointers move-behavior.
     * <P>
     * If the horizontalStepDistance is set to a negative value
     * the pointer will move to the left when movement is set to
     * MOVE_RIGHT and to the right when movement is set to MOVE_LEFT.
     * <P>
     * If the horizontalStepDistance is set to <code>0</code>,
     * the pointer will not move when movement is set to 
     * MOVE_RIGHT or MOVE_LEFT.
     * 
     * @param distance	The number of cells to move in the horizontal
     * 					axis.
     * @see #setMovement(int)
     */
    public void setHorizontalStepDistance(int distance) {
        hStep = distance;
    }
    
    /**
     * Gets the horizontalStepDistance.
     */
    public int getHorizontalStepDistance() {
        return hStep;
    }
    
    /**
     * Sets the number of cells this cellpointer will move 
     * in the vertical axis after a call to {@link #move()}.
     * The default is 1.
     * <P>
     * If the movement of this cellpointer is set to
     * MOVE_RIGHT or MOVE_LEFT the value of verticalStepDistance
     * has no influence on this pointers move-behavior.
     * <P>
     * If the verticalStepDistance is set to a negative value
     * the pointer will move up when movement is set to
     * MOVE_DOWN and down when movement is set to MOVE_UP.
     * <P>
     * If the verticalStepDistance is set to <code>0</code>,
     * the pointer will not move when movement is set to 
     * MOVE_DOWN or MOVE_UP.
     * 
     * @param distance	The number of cells to move in the vertical
     * 					axis.
     * @see #setMovement(int)
     */
    public void setVerticalStepDistance(int distance) {
        vStep = distance;
    }
    
    /**
     * Gets the verticalStepDistance.
     */
    public int getVerticalStepDistance() {
        return vStep;
    }
    
    /**
     * Sets the direction this cellpointer will move after a call to
     * {@link #move()}. The default is MOVE_RIGHT.
     * 
     * @param moveConst	One of CellPointer's MOVE_RIGHT, MOVE_LEFT, MOVE_DOWN
     * 					or MOVE_UP values.
     * @throws IllegalArgumentException at values less than 0 or greater then 3.
     */
    public void setMovement(int moveConst) {
        switch (moveConst) {
        	case MOVE_RIGHT: hMove = 1; vMove = 0; break;
        	case MOVE_LEFT: hMove = -1; vMove = 0; break;
        	case MOVE_DOWN: hMove = 0; vMove = 1; break;
        	case MOVE_UP: hMove = 0; vMove = -1; break;
        	default: throw new IllegalArgumentException(
        	        moveConst + ". Legal values are 0, 1, 2 and 3.");
        }
    }
    
    /**
     * Gets the direction into which this cellpointer will move. The returned int will
     * be equal to one of CellPointer's MOVE_RIGHT, MOVE_LEFT, MOVE_DOWN
     * or MOVE_UP values.
     */
    public int getMovement() {
        if (hMove == 1) return MOVE_RIGHT;
        if (hMove == -1) return MOVE_LEFT;
        if (vMove == 1) return MOVE_DOWN;
        return MOVE_UP;
    }
    
    /**
     * Moves this cellpointer. The direction of the movement depends on the 
     * setting of {@link #setMovement(int)}. How many cells the pointer will
     * move depends on the setting of the step distance.
     * 
     * @see #setMovement(int)
     * @see #setHorizontalStepDistance(int)
     * @see #setVerticalStepDistance(int)
     */
    public void move() {
        c += (hStep * hMove);
        r += (vStep * vMove);
    }
    
    /**
     * Moves this cellpointer to a new position relative to it's old position.
     * If this cellpointer was previously at oldR, oldC, the new position will 
     * be at oldR + rows, oldC + columns.
     * 
     * @param rows the number of rows to move
     * @param columns the number of columns to move
     */
    public void move(int rows, int columns) {
        r += rows;
        c += columns;
    }
    
    /**
     * Moves this cellpointer to the specified row and column number.
     * 
     * @param row the row where this cellpointer should move to
     * @param column the column where this cellpointer should move to
     */
    public void moveTo(int row, int column) {
        r = row;
        c = column;
    }
    
    /**
     * Moves this cellpointer to the address specified by the given A1-reference
     * string.
     * 
     * @param a1_ref a string of A1-reference type
     */
    public void moveTo(String a1_ref) {
        r = calculateRow(a1_ref);
        c = calculateColumn(a1_ref);
    }
    
    /**
     * Moves this cellpointer to the specified address.
     * 
     * @param address the address where this cellpointer should move to
     */
    public void moveTo(Address address) {
        r = address.r;
        c = address.c;
    }
    
    /**
     * Moves this cellpointer to the first column of the present row.
     * The first column is determined by the value of 
     * {@link nl.fountain.xelem.excel.Worksheet#firstColumn}.
     */
    public void moveHome() {
        c = Worksheet.firstColumn;
    }
    
    /**
     * Moves this cellpointer to the first column of the next row.
     * The first column is determined by the value of 
     * {@link nl.fountain.xelem.excel.Worksheet#firstColumn}.
     * How many rows the new position will be from the present position
     * is determined by the value of 
     * {@link #setVerticalStepDistance(int) verticalStepDistance}.
     *
     */
    public void moveCRLF() {
        r += vStep;
        c = Worksheet.firstColumn;
    }

}

package com.veinhorn.xelem.test;

import nl.fountain.xelem.XSerializer;
import nl.fountain.xelem.XelemException;
import nl.fountain.xelem.excel.Cell;
import nl.fountain.xelem.excel.Row;
import nl.fountain.xelem.excel.Workbook;
import nl.fountain.xelem.excel.Worksheet;

/**
 * Created by Boris Korogvich on 02.08.2017.
 * Useful DSL for sheet creation
 */
public abstract class XelemBuilder {

    public static XelemBuilder with(Workbook workbook) {
        return new XelemBuilderImpl(workbook);
    }

    abstract SheetBuilder newSheet();
    // abstract XelemBuil

    abstract Workbook build();
    abstract void serialize() throws XelemException;

    interface Builder<T> {
        T build();
    }

    interface SheetBuilder extends Builder<XelemBuilder> {
        RowBuilder addRow();
    }

    interface RowBuilder extends Builder<SheetBuilder> {
        CellBuilder addCell();
    }

    interface CellBuilder extends Builder<RowBuilder> {
        RowBuilder data(String data);
    }


    private static class XelemBuilderImpl extends XelemBuilder {
        private Workbook workbook;

        public XelemBuilderImpl(Workbook workbook) {
            this.workbook = workbook;
        }

        @Override
        public SheetBuilder newSheet() {
            return new SheetBuilderImpl(this, workbook.addSheet());
        }

        @Override
        public Workbook build() {
            return workbook;
        }

        @Override
        public void serialize() throws XelemException {
            new XSerializer().serialize(workbook);
        }
    }

    private static class SheetBuilderImpl implements SheetBuilder {
        protected XelemBuilder builder;
        private Worksheet worksheet;

        public SheetBuilderImpl(XelemBuilder builder, Worksheet worksheet) {
            this.builder = builder;
            this.worksheet = worksheet;
        }

        @Override
        public RowBuilder addRow() {
            return new RowBuilderImpl(this, worksheet.addRow());
        }

        @Override
        public XelemBuilder build() {
            return builder;
        }
    }

    private static class RowBuilderImpl implements RowBuilder {
        private SheetBuilder builder;
        private Row row;

        public RowBuilderImpl(SheetBuilder builder, Row row) {
            this.builder = builder;
            this.row = row;
        }

        @Override
        public CellBuilder addCell() {
            return new CellBuilderImpl(this, row.addCell());
        }

        @Override
        public SheetBuilder build() {
            return builder;
        }
    }

    private static class CellBuilderImpl implements CellBuilder {
        private RowBuilder builder;
        private Cell cell;

        public CellBuilderImpl(RowBuilderImpl builder, Cell cell) {
            this.builder = builder;
            this.cell = cell;
        }

        @Override
        public RowBuilder data(String data) {
            cell.setData(data);
            return builder;
        }

        @Override
        public RowBuilder build() {
            return builder;
        }
    }
}

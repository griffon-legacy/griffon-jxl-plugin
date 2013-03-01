/*
 * Copyright 2012-2013 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package griffon.plugins.jxl

import jxl.write.*
import jxl.write.WritableFont.*

/**
 * @author Shaun Jurgemeyer
 */
class ExcelUtils {

    void copyCellFormatWithValue(sheet, origCol, origRow, newCol, newRow, newValue) {

        def format = sheet.getCell(origCol,origRow).cellFormat
        Cell cell = new Cell(newCol, newRow, newValue)
        if (format) {
            cell.format = new WritableCellFormat(format)
            cell.font = new WritableFont(format.font)
        }
        cell.write(sheet)
    }

    void copyDown(sheet, origCol, origRow, newRow) {
        sheet.addCell(sheet.getCell(origCol,origRow).copyTo(origCol, newRow))
    }

    void copyDown(sheet, origCol, origRow, newRow, newValue) {
        copyCellFormatWithValue(sheet, origCol, origRow, origCol, newRow, newValue)
    }

    void mergeAcross(sheet, startCol, endCol, row) {
       sheet.mergeCells(startCol, row, endCol, row)
    }

    void withTemplateRow(sheet, templateRow, items, Closure closure) {
        def row = templateRow+1
        items.each {
            sheet.insertRow(row)
            closure(it, templateRow, row)
            row++
        }
        sheet.removeRow(templateRow)
    }

    Cell getCell(sheet, int col, int row, Map props=[:]) {
        def jxlCell = sheet.getCell(col,row)
        new Cell(jxlCell, props)
    }

    private void eachCell(sheet,cellList,Closure closure) {
        cellList.each {
            def cell = getCell(sheet, it[0], it[1])
            closure(cell)
            cell.write(sheet)
        }
    }

    void addData(sheet,rowData,startCol=0,startRow=0) {
        rowData.eachWithIndex { row, rowNumber ->
            row.eachWithIndex { col, colNumber ->
                new Cell(startCol+colNumber, startRow+rowNumber, col).write(sheet)
            }
        }
    }

    def createWorkbook(String filePath) {
        jxl.Workbook.createWorkbook(new File(filePath))
    }

    def createWorkbook(File file) {
        jxl.Workbook.createWorkbook(file)
    }

    def createWorkbook(OutputStream stream) {
        jxl.Workbook.createWorkbook(stream)
    }
}

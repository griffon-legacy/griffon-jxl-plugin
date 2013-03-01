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

package griffon.plugins.jxl.builder

import griffon.plugins.jxl.*

/**
 * @author Shaun Jurgemeyer
 */
@Mixin(ExcelUtils)
class ExcelBuilder {

    def workbook
    def sheet
    def sheetIndex = 0
    def formula = new ExcelFormulaBuilder()
    def cells = []

    def workbook(String fileName, Closure closure) {
        closure.resolveStrategy =  Closure.DELEGATE_FIRST
        closure.delegate = this
        this.workbook = createWorkbook(fileName)
        sheetIndex = 0
        closure()
        workbook.write()
        workbook.close()
    }

    def workbook(OutputStream stream, Closure closure) {
        closure.resolveStrategy =  Closure.DELEGATE_FIRST
        closure.delegate = this
        this.workbook = createWorkbook(stream)
        sheetIndex = 0
        closure()
        workbook.write()
        workbook.close()
    }

    def sheet(String name="Sheet$sheetIndex", Closure closure) {
        this.sheet = workbook.createSheet(name, sheetIndex++)
        this.cells = []
        closure()
        cells.each {
            it.write(sheet)
        }
    }

    def cell(int col, int row, value, Map props=[:]) {
        def newCell = new Cell(col, row, value, props)
        cells << newCell
        newCell
    }

    def cell(int col, int row, Map props=[:]) {
        def newCell = getCell(sheet, col, row, props)
        cells << newCell
        newCell
    }

    def addData(rowData,startCol=0,startRow=0) {
        addData(sheet, rowData, startCol, startRow)
    }
}

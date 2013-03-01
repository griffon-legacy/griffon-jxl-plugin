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

/**
 * @author Shaun Jurgemeyer
 */
class ExcelFormulaBuilder {
    private static final String letters='ABCDEFGHIJKLMNOPQRSTUVWXYZ'

    String cell(int colNum, int rowNum) {
        def col = numberToLetter(colNum)
        "$col${rowNum+1}"
    }

    String range(int colStart, int rowStart, int colEnd, int rowEnd) {
        "${cell(colStart, rowStart)}:${cell(colEnd, rowEnd)}"
    }

    String numberToLetter(int num) {
       num++
       String result = ''
       while (num > 0) {
           int remainder = num % 26
           result = letters[remainder-1] + result
           num = (num-1)/26
       }
       result
    }

    def methodMissing(String name, args) {
        "=${name.toUpperCase()}(${args.join(',')})"
    }
}

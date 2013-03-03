/*
 * Copyright 2013 the original author or authors.
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

/**
 * @author Andres Almiray
 */
class JxlGriffonPlugin {
    // the plugin version
    String version = '0.2'
    // the version or versions of Griffon the plugin is designed for
    String griffonVersion = '1.2.0 > *'
    // the other plugins this plugin depends on
    Map dependsOn = [:]
    // resources that are included in plugin packaging
    List pluginIncludes = []
    // the plugin license
    String license = 'Apache Software License 2.0'
    // Toolkit compatibility. No value means compatible with all
    // Valid values are: swing, javafx, swt, pivot, gtk
    List toolkits = []
    // Platform compatibility. No value means compatible with all
    // Valid values are:
    // linux, linux64, windows, windows64, macosx, macosx64, solaris
    List platforms = []
    // URL where documentation can be found
    String documentation = ''
    // URL where source can be found
    String source = 'https://github.com/griffon/griffon-jxl-plugin'

    List authors = [
        [
            name: 'Shaun Jurgemeyer',
            email: 'sjurgemeyer@gmail.com'
        ],
        [
            name: 'Andres Almiray',
            email: 'aalmiray@yahoo.com'
        ]
    ]
    String title = 'Exports data to Excel using the JXL library'
    // accepts Markdown syntax. See http://daringfireball.net/projects/markdown/ for details
    String description = '''
This plugin allows you to easily create formatted Excel documents with a simple
Builder-like syntax. It utilizes the [JXL][] Library which is the most full featured
Java library for interacting with Excel. Thanks need to go to Andy Khan for creating
this great library. This is a port of the Grails [jxl plugin][] original by
Shaun Jurgemeyer.

Usage
-----

### Basic Worksheet Creation ### 

You can then use the ExcelBuilder class to create your worksheet by providing a
file path. Within that block you can then declare sheets and cells.

    new ExcelBuilder().workbook('/path/to/test.xls') {
        sheet('SheetName') {
            cell(0,0,'DEF')
        }
    }

The cell method takes three arguments: column, row and value and returns a
`griffon.plugins.jxl.Cell` object. Value can be a String, Number or Date. The type
of object determines the type of cell that will be created in Excel.

Notice that cells are zero indexed. This is a bit confusing when working with Excel
rows, but is consistent with how JXL works and makes coding a little easier.

### Formatting ###

There are many for formatting your cells. The JXL plugin adds several methods to
the cell object by default. For example:

    cell(0,9,'Tenth Cell').bold().italic().center()

Creates a cell that is bold, italic and centered. Notice that these methods all
return the cell back so they can be chained.

The built in methods are:

    //Font style
    bold()
    italic()
    
    //Alignment
    center()
    centre() //alternate spelling based on JXL Library
    left()
    right()
    justify()
    fill() 
    
    //Borders
    dashDotBorder() 
    dashDotDotBorder() 
    dashedBorder() 
    dottedBorder() 
    doubleBorder() 
    hairBorder() 
    mediumBorder() 
    mediumDashDotBorder() 
    mediumDashDotDotBorder() 
    mediumDashedBorder() 
    noneBorder() 
    slantedDashDotBorder() 
    thickBorder() 
    thinBorder() 
    
    wrap() //Wrap text

### Formulas ###

You can generate Excel formulas using dynamic methods as well. For example:

    cell(3,6, formula.sum(formula.range(3,0,3,5)))

creates a cell at cell G4 with a formula of `=SUM(A4:F4)`

The formula object has a range function to produce the `"A4:F4"` using cell
indices. It also has a dynamic method for any Excel function, in this case `SUM`.
Basically any function called on formula will produce a call to that Excel function
with the provided parameters.

There is a current issue with this as literal Strings are not quoted as they should
be. In order to get around this, for now you will have to pre-quote strings.

### Additional JXL functionality. ###

If you are familiar with the JXL library, you know that there is much more that
you can do with the format etc. While there is not currently any special support
for these other features, you can still access them. The `griffon.plugins.jxl.Cell`
passes through all method calls to WritableFont and WritableFormat as necessary.
You can view the api docs for JXL [here][].


[JXL]: http://jexcelapi.sourceforge.net/
[jxl plugin]: http://grails.org/plugin/jxl
[here]: http://jexcelapi.sourceforge.net/resources/javadocs/2_6_10/docs/index.html
'''
}

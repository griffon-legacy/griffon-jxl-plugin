
Exports data to Excel using the JXL library
-------------------------------------------

Plugin page: [http://artifacts.griffon-framework.org/plugin/jxl](http://artifacts.griffon-framework.org/plugin/jxl)


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


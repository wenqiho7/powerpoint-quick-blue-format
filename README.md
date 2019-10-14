# powerpoint-quick-blue-format
Used in the organisation that I served an internship with. This VBA code snippet offers a standard PowerPoint colour palette as well as a standardised (bar/line/pie) chart format, size,  legend, labels and annotation which are all quickly applied with one click.

## Background
At the company where I served an internship with, there are always slides to be made. Slides that have to be formatted to the standard color palette, the font, the size, the labels and annotations. The amount of time to format just one chart and the effort to ensure standardisation across the slide deck served as motivation for this project, a one-click solution to standardised slide charts

## Directions of use
This VBA code was written on PowerPoint VBA. Copy and paste the code into a new module on PowerPoint VBA.
For Windows users: select 'Customize the Ribbon…' and in the window that appears, tick the Developer Tab check box on the right hand side and click OK to close the window. Click on the Visual Basic button within this tab > Insert > Module
For Mac users (I turned into one recently): select Tools > Macro > Visual Basic Editor > Insert > Module

After inserting the code, go to the View ribbon tab and click Macros. You should see a list of 5 macros that you have added: additions, all_barline, all_pie, barline, pie

### additions
To use: click on a chart shape, then click on the macro
What it does: inserts a bubble to input CAGR and inserts a textbox to indicate the units for y-axis

### all_barline
To use: on the preferred slide with bar/line charts, click on the macro
What it does: all bar/line charts on that slide will be formatted to 6 standardised shades of blue, along with legend placement, font and font size

### all_pie
To use: on the preferred slide with pie charts, click on the macro
What it does: all pie charts on that slide will be formatted to 6 standardised shades of blue, along with legend placement, font and font size

### barline
To use: click on a bar/line chart shape, then click on the macro
What it does: selected bar/line chart will be formatted to 6 standardised shades of blue, along with legend placement, font and font size

### pie
To use: click on a pie chart shape, then click on the macro
What it does: selected pie chart will be formatted to 6 standardised shades of blue, along with legend placement, font and font size

## Note
All macros in this code allow for any bar/line/pie chart but only limited to these chart types.
The code has standardised set of blue colors for charts with up to 6 series.
The code is written for my use during my internship at the organisation, hence the organisation's blue shades was used. Of course, feel free to change the colors, fonts, sizes and annotations!

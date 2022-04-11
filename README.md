This is all about developing Visio solutions.

- Available technologies and which one to chose
  - Add-Ons    
  [Save As addon for Visio VBA](https://www.paulherber.co.uk/free-visio-addons/#saveas) by [Paul Herber](https://github.com/paulherber)     
  *Addon to allow Visio VBA to do a "Save As"
  There is no built-in dialog in Visio VBA to do a Save As. So, I have created an addon that can do this.
Download the .zip file below, unzip it and run the .msi file contained within. This will install the addon in Visio.
Now download the test.zip file below and unzip it into your My Shapes folder. Open this file and look at the macro "SaveAs".
Edit the contents of the string "SaveAddress" to change the default location and name where the file is saved to.*    
  - Plug-Ins    
  [VBA Import/Export Addin](https://unmanagedvisio.com/products/vba-importexport-addin/) by [Nikolay Belykh](https://github.com/nbelyh)     
  *This free Visio add-in helps to import and export VBA code from Visio to a folder.
  The addin can exports and import the code from “ThisDocument”, forms (.frm), classes (.cls) and modules (.bas)*    
  [ShapeSheet Watch](https://unmanagedvisio.com/shapesheet-watch-multiselect/) by [Nikolay Belykh](https://github.com/nbelyh)     
  *ShapeSheet Watch is an additional window for easy work with ShapeSheet cells in MS Visio.*    
  [Visio Power Tools2010](https://viziblr.com/news/2012/3/10/browsing-visio-2010-stencil-shapes-as-a-document.html) by [Saveen Reddy](https://github.com/saveenr)    
  *Add-in for create an HTML document that contains lists and thumbnails of all the Masters in a Visio Stencil.*
  - Interactions with other office applications    
    [Visio BackSync](https://unmanagedvisio.com/products/visio-back-sync/) by [Nikolay Belykh](https://github.com/nbelyh)     
    *This add-in uses Visio data-binding infrastructure, and allows you to update the data in the external source from Visio (back-synchronize). There are limitations of course, but basically it supports any type of data source which support update (including Excel files, databases, and SharePoint lists).*
  - Templates, Stencils and macros
  - Smartshapes


- Where to start...
- Developing    
  - [Extended Visio Addin Project](https://github.com/nbelyh/VisioPanelAddinVSTO#extended-visio-addin-project) by [Nikolay Belykh](https://github.com/nbelyh)     
  *This extension contains a project template to create a extneded Microsoft Visio Addin based on Visual Tools for Office or vanilla COM Shared AddIn in C# and VB.NET. [Release 1.2.2](https://github.com/nbelyh/VisioPanelAddinVSTO/releases/download/1.2.2/VisioPanelVSTOAddinVSIX_2019.vsix)*    
  - [IdMsoAutocomplete](https://github.com/nbelyh/IdMsoAutocomplete#idmsoautocomplete) by [Nikolay Belykh](https://github.com/nbelyh)     
  *The extension provides autocomplete for CustomUI for office ribbon (means, auto-complete for the built-in item identifiers, like idMso, insertBeforeMso, insertAfterMso and imageMso). [Release 1.0.3](https://github.com/nbelyh/IdMsoAutocomplete/releases/download/1.0.3/IdMsoAutocomplete.2019.vsix)*
- Testing
- Deploying
- Maintaining

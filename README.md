# dita-excel
DITA Open Toolkit plugin which dynamically converts to DITA Excel files referenced with format="excel" in DITA Maps.

The converted DITA topic has as ID the file name without extension of the Excel document and inside contains a DITA table for each sheet, with the table ID being equal to the sheet name.

The "samples" folder contains a sample DITA Map showing how a topic can have a conref to a table directly from the Excel document.

The plugin should work with both DITA Open Toolkit 2.5.4 and 3.2.1. To install the plugin in a DITA Open Toolkit:

    - Download the entire project as a ZIP file.
    - Unzip and copy the "com.oxygenxml.excel.dita" folder to the "DITA-OT-DIR\plugins" folder.
    - Run the DITA OT integrator (https://www.dita-ot.org/dev/topics/plugins-installing.html)

Note: If you are using Oxygen, you can run the DITA OT Integrator by following this procedure: https://www.oxygenxml.com/doc/ug-editor/topics/dita-ot-install-plugin.html

Copyright and License
---------------------
Copyright 2019 Syncro Soft SRL.

This project is licensed under [Apache License 2.0](https://github.com/oxygenxml/dita-excel/blob/master/LICENSE)


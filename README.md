# dita-excel
Dynamically converts to DITA Excel files referenced with format="excel" in DITA Maps.

The converted DITA topic has as ID the file name without extension of the Excel document and inside contains a DITA table for each sheet, with the table ID being equal to the sheet name.

The "samples" folder contains a sample DITA Map showing how a topic can have a conref to a table directly from the Excel document.

Copyright and License
---------------------
Copyright 2018 Syncro Soft SRL.

This project is licensed under [Apache License 2.0](https://github.com/oxygenxml/dita-excel/blob/master/LICENSE)


/// <reference path="Scripts/_references.js" />

var tableXmlBinder = (function () {
    'use strict';

    var tableName = "PersonsTable";
    var tableNameFull = "Sheet2!PersonsTable";
    var bindingName = "PersonsBinding";
    var binding = null;
    var xmlPartDoc = null;

    //
    // Initializes the Custom XML for this sample and binds to the Excel table for change notifications
    //
    function initTableContentXmlBinding() {

        console.log("Start table xml binding...");

        // Load the XML data
        Excel.run(function (context) {

            console.log("Loading XML part...");

            var parts = context.workbook.customXmlParts;
            parts.load();

            var xmlPart = parts.getByNamespace("SapPrototypeTest").getOnlyItem();
            xmlPart.load();

            return context.sync().then(function () {
                // No exception occured, so the XML part is present
                console.log("Xml part found, loading XML content...");
                var xmlData = xmlPart.getXml();
                return context.sync().then(function () {
                    console.log("Successfully loaded XML content, now loading content into DOM!");
                    xmlPartDoc = $.parseXML(xmlData.value);
                });
            }).catch(function (error) {
                // Exception occured for xmlPart.load(), so must not be present
                console.log("Failed loading existing XML part, adding new one...");
                var newXml = '<?xml version="1.0"?><PersonsData xmlns="SapPrototypeTest"></PersonsData>';
                xmlPart = parts.add(newXml);
                xmlPart.load();
                return context.sync().then(function () {
                    console.log("Added new XML part, now loading XML document...");
                    var xmlData = xmlPart.getXml();
                    return context.sync().then(function () {
                        console.log("Loading XML Document into DOM...");
                        xmlPartDoc = $.parseXML(xmlData.value);
                        return context.sync();
                    });
                });
            });

        }).then(function () {

            console.log("XML part available, binding table...");

            // Start binding to the the table
            Office.context.document.bindings.addFromNamedItemAsync(
                tableNameFull,
                Office.BindingType.Table,
                { id: bindingName },
                function (results) {
                    binding = results.value;
                    addTableBindingHandler();
                }
            );

        }).catch(function (error) {
            console.error(">>>>>>>>>>>>>>>> ERROR >>>>>>>>>>>>>>>>>>");
            console.error(error);
        });

    }

    function addTableBindingHandler(callback) {
        Office.select("bindings#" + bindingName).addHandlerAsync(
            Office.EventType.BindingDataChanged,
            onBindingDataChanged,
            function () {
                if (callback) { callback(); }
            }
        );
    }

    // Called when data in the table changes
    var onBindingDataChanged = function (result) {
        console.log('mszcool binding data changed');
    };


    //
    // Saves the content to the XML island in Excel
    // 
    function saveTableContentToXml() {

        Excel.run(function (context) {

            var theTable = context.workbook.tables.getItem(tableName);
            var headerRange = theTable.getHeaderRowRange();
            headerRange.load();

            var parts = context.workbook.customXmlParts;
            parts.load();

            var xmlPart = parts.getByNamespace("SapPrototypeTest").getOnlyItem();
            xmlPart.delete();

            return context.sync().then(function () {

                binding.getDataAsync(
                    {
                        startRow: 0,
                        startColumn: 0
                    },
                    function (asyncResult) {
                        var valuesInTable = asyncResult.value;
                        saveRowsInXml(headerRange, valuesInTable, xmlPartDoc);

                        var xmlString = (new XMLSerializer()).serializeToString(xmlPartDoc);

                        xmlPart = parts.add(xmlString);
                        xmlPart.load();

                        return context.sync().then(function () {
                            var xmlData = xmlPart.getXml();
                            return context.sync().then(function () {
                                xmlPartDoc = $.parseXML(xmlData.value);
                            });
                        });
                    }
                );

            });
        }).catch(function (error) {
            console.error(">>>>>>>>>>>>>> Error when saving XML >>>>>>>>>>>>>>>>");
            console.error(error);
        });

    }

    //
    // Saves each row in the bound table to the XML data
    //
    function saveRowsInXml(tableHeader, valuesInTable, xmlPartDoc) {

        // First, get the PersonsData Element from the XML
        var personsNode = xmlPartDoc.getElementsByTagName('PersonsData')[0];

        // Walk through each row in the table bound
        for (var i = 0; i < valuesInTable.rows.length; i++) {

            // Now try to find the existing row in the PersonsData XML
            var isNewPerson = false;
            var personElement = getElementByIdXml(xmlPartDoc, valuesInTable.rows[i][0].toString());
            if (!personElement) {
                // Person does not exist, create a new element
                isNewPerson = true;
                personElement = xmlPartDoc.createElement("Person");
                personElement.setAttribute("ID", valuesInTable.rows[i][0]);
            }

            // Now set the values for each column in the Excel Table
            for (var col = 0; col < tableHeader.values[0].length; col++) {
                if (isNewPerson) {
                    var personColumnElement = xmlPartDoc.createElement(tableHeader.values[0][col].toLowerCase());
                    personColumnElement.appendChild(xmlPartDoc.createTextNode(valuesInTable.rows[i][col]));
                    personElement.appendChild(personColumnElement);
                } else {
                    var personColumnElements = personElement.getElementsByTagName(tableHeader.values[0][col].toLowerCase());
                    if (personColumnElements.length > 0) {
                        personColumnElements[0].childNodes[0].nodeValue = valuesInTable.rows[i][col];
                    } else {
                        // Column was added while last save... so add a new column
                        var newPersonColumnElement = xmlPartDoc.createElement(tableHeader.values[0][col].toLowerCase());
                        newPersonColumnElement.appendChild(xmlPartDoc.createTextNode(valuesInTable.rows[i][col]));
                        personElement.appendChild(newPersonColumnElement);
                    }
                }
            }

            // Finally append the new person if it is indeed a new one
            if (isNewPerson) {
                personsNode.appendChild(personElement);
            }
        }

    }

    // 
    // Helper to make getElementById() work with XML without a DTD
    //
    function getElementByIdXml(the_node, the_id) {
        //get all the tags in the doc
        var node_tags = the_node.getElementsByTagName('Person');
        for (var i = 0; i < node_tags.length; i++) {
            //is there an id attribute?
            if (node_tags[i].hasAttribute('id')) {
                //if there is, test its value
                if (node_tags[i].getAttribute('id') === the_id) {
                    //and return it if it matches
                    return node_tags[i];
                }
            } else if (node_tags[i].hasAttribute('ID')) {
                if (node_tags[i].getAttribute('ID') === the_id) {
                    //and return it if it matches
                    return node_tags[i];
                }
            }
        }
        // Nothing found
        return null;
    }


    //
    // The signature return from this prototype
    // 
    return {
        startBinding: initTableContentXmlBinding,
        saveContent: saveTableContentToXml
    };
})();
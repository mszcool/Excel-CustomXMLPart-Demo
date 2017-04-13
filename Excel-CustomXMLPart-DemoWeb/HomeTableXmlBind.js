/// <reference path="Scripts/_references.js" />

var tableXmlBinder = (function () {
    'use strict';

    var tableName = "Sheet2!PersonsTable";
    var bindingName = "PersonsBinding";
    var binding = null;
    var xmlPart = null;
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
                tableName,
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

    };

    function addTableBindingHandler(callback) {
        Office.select("bindings#" + bindingName).addHandlerAsync(
            Office.EventType.BindingDataChanged,
            onBindingDataChanged,
            function () {
                if (callback) { callback(); }
            }
        );
    };

    // Called when data in the table changes
    var onBindingDataChanged = function (result) {
        console.log('mszcool binding data changed');
    };


    //
    // Saves the content to the XML island in Excel
    // 
    function saveTableContentToXml() {

        // Get the Excel Table and walk through the lines
        Excel.run(function (context) {

            var bindings = context.workbook.bindings(tableName);

        }).catch(function (error) {
            console.error(">>>>>>>>>>>>>>>>>> ERROR Saving XML >>>>>>>>>>>>>>>>");
            console.error(error);
        });

    }


    //
    // The signature return from this prototype
    // 
    return {
        startBinding: initTableContentXmlBinding,
        saveContent: saveTableContentToXml
    };
})();
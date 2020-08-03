Office.initialize = function (reason) {

    var counter = 1;
    var itemId = "";
    var result = "";

    var xpathsCustomXMLPartId = "";
    var conditionsCustomXMLPartId = "";

    const namespaceEnum = {
        XPATHS: 'http://opendope.org/xpaths',
        CONDITIONS: 'http://opendope.org/conditions'
    }

    const customXMLPart = {
        XPATHS: '<?xml version="1.0" encoding="UTF-8"?><xpaths xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns="http://opendope.org/xpaths"></xpaths>',
        CONDITIONS: '<?xml version="1.0" encoding="UTF-8"?><conditions xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://opendope.org/conditions"></conditions>'
    }

    const coercicionEnum = {
        TEXT: 'text',
        HTML: 'html',
        IMAGE: 'image',
        CONDITION: 'condition', 
        REPEAT: 'repeat'
    }

    function checkIfCustomXMLPartsExist() {
        checkIfCustomXMLPartExists(namespaceEnum.XPATHS, customXMLPart.XPATHS);
        checkIfCustomXMLPartExists(namespaceEnum.CONDITIONS, customXMLPart.CONDITIONS);
    }

    function checkIfCustomXMLPartExists(namespace, body) {
        Office.context.document.customXmlParts.getByNamespaceAsync(
            namespace,
            function (result) {
                var xmlPart = result.value;
                if (xmlPart.length === 0) {
                    Office.context.document.customXmlParts.addAsync(
                        body,
                        function (result) {
                            if (namespace == namespaceEnum.XPATHS)
                                xpathsCustomXMLPartId = result.value.id;
                            else if (namespace == namespaceEnum.CONDITIONS)
                                conditionsCustomXMLPartId = result.value.id;
                        }
                    );
                }
                else {
                    if (namespace == namespaceEnum.XPATHS)
                        xpathsCustomXMLPartId = xmlPart[0].id;
                    else if (namespace == namespaceEnum.CONDITIONS)
                        conditionsCustomXMLPartId = xmlPart[0].id;
                }
            }
        );
    }

    function generateDataDefinitionTree() {
        Office.context.document.customXmlParts.getByNamespaceAsync("http://opendope.org/datasourceDefinition", function (result) {
            var xmlPart = result.value;
            if (xmlPart.length !== 0) {
                itemId = xmlPart[0].id;

                Office.context.document.customXmlParts.getByIdAsync(itemId, function (result) {
                    var xmlPart = result.value;
                    xmlPart.getXmlAsync(function (result) {
                        if (result.status !== "succeeded") {
                            return;
                        }

                        var xml = $.parseXML(result.value);
                        var json = mapXMLToJSTreeFormat(xml, "#");
                        json = json.substring(0, json.length - 1);

                        var body = '{"core": {"data": [TREE]}}';
                        var data = body.replace("TREE", json);

                        $('#tree').jstree($.parseJSON(data));
                        createMenu();
                    });
                });
            }
        });
    }

    function createMenu() {
        $("#tree").on("select_node.jstree", function (e, data) {
            var node = data.node.text;
            var nodeIsFolder = data.node.children.length != 0;
            var xpath = "/" + $('#tree').jstree("get_path", data.node.id, "[1]/") + (!nodeIsFolder ? "[1]" : "");

            if (!nodeIsFolder) {
                var menu = [{
                    name: 'Insérer un contrôle de contenu',
                    title: 'Insérer un contrôle de contenu',
                    fun: function () {
                        createContentControl(node, xpath, coercicionEnum.TEXT);
                    }
                }, {
                    name: 'Insérer un contrôle de contenu de type image',
                    title: 'Insérer un contrôle de contenu de type image',
                    fun: function () {
                        createContentControl(node, xpath, coercicionEnum.IMAGE);
                    }
                }, {
                    name: 'Insérer un contrôle de contenu de type HTML',
                    title: 'Insérer un contrôle de contenu de type HTML',
                    fun: function () {
                        createContentControl(node, xpath, coercicionEnum.HTML);
                    }
                }, {
                    name: 'Insérer un contrôle de contenu de type condition',
                    title: 'Insérer un contrôle de contenu de type condition',
                    fun: function () {
                        createContentControl(node, xpath, coercicionEnum.CONDITION);
                    }
                }];
            }
            else {
                var menu = [{
                    name: 'Insérer un contrôle de contenu de type repeat',
                    title: 'Insérer un contrôle de contenu de type repeat',
                    fun: function () {
                        createContentControl(node, xpath, coercicionEnum.REPEAT);
                    }
                }];
            }

            $(data.event.target).contextMenu(
                menu,
                {
                    triggerOn: 'click',
                    mouseClick: 'right',
                    position: 'left',
                    top: 'auto',
                    left: 10
                }
            );
        });
    }

    function mapXMLToJSTreeFormat(entry, parent) {
        for (var i = 0; i < entry.childNodes.length; i++) {
            var node = entry.childNodes[i];
            var id = counter;
            counter++;

            if (node.childNodes.length !== 0 && node.childNodes[0].nodeType !== 3) {
                result += '{"id": "' + id + '", "parent": "' + parent + '", "text": "' + node.nodeName + '" },';
                mapXMLToJSTreeFormat(node, id);
            } else {
                result += '{"id": "' + id + '", "parent": "' + parent + '", "text": "' + node.nodeName + '", "icon": "jstree-file" },';
            }
        }

        return result;
    }

    function generateContentControlUniqueId(node) {
        var id = "";
        var characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz";
        for (var i = 0; i < 5; i++) {
            id += characters.charAt(Math.floor(Math.random() * characters.length));
        }

        return node + "_" + id;
    }

    function getContentControlBodyTemplate(type) {
        if (type == coercicionEnum.TEXT) {
            return '<w:sdt><w:sdtPr><w:alias w:val="Data value: {{ccID}}"/><w:tag w:val="od:xpath={{ccID}}"/><w:id w:val="1330132901"/></w:sdtPr><w:sdtContent><w:proofErr w:type="spellStart"/><w:r><w:t>{{text}}</w:t></w:r><w:proofErr w:type="spellEnd"/></w:sdtContent></w:sdt>';
        }
        else if (type == coercicionEnum.HTML) {
            return '<w:sdt><w:sdtPr><w:alias w:val="XHTML: {{ccID}}"/><w:tag w:val="od:xpath={{ccID}}&amp;od:ContentType=application/xhtml+xml"/><w:id w:val="2110929613"/></w:sdtPr><w:sdtContent><w:proofErr w:type="spellStart"/><w:r><w:t>{{text}}</w:t></w:r><w:proofErr w:type="spellEnd"/></w:sdtContent></w:sdt>';
        }
        else if (type == coercicionEnum.IMAGE) {
            return '<w:sdt><w:sdtPr><w:alias w:val="Data value: {{ccID}}"/><w:tag w:val="od:xpath={{ccID}}"/><w:id w:val="1330132901"/></w:sdtPr><w:sdtContent><w:proofErr w:type="spellStart"/><w:r><w:t>{{text}}</w:t></w:r><w:proofErr w:type="spellEnd"/></w:sdtContent></w:sdt>';
        }
        else if (type == coercicionEnum.CONDITION) {
            return '<w:sdt><w:sdtPr><w:alias w:val="Conditional: {{ccID}}"/><w:tag w:val="od:condition={{ccID}}"/><w:id w:val="1801413517"/></w:sdtPr><w:sdtContent><w:r><w:t xml:space="preserve"> </w:t></w:r></w:sdtContent></w:sdt>';
        }
        else if (type == coercicionEnum.REPEAT) {
            return '<w:sdt><w:sdtPr><w:alias w:val="Repeat: {{ccID}}"/><w:tag w:val="od:repeat={{ccID}}"/><w:id w:val="-628546810"/></w:sdtPr><w:sdtContent><w:r><w:t xml:space="preserve"> </w:t></w:r></w:sdtContent></w:sdt>';
        }
    }

    function getContentControlBody(template, id, node, xpath) {
        var body = template;
        body = body.replace("{{ccID}}", id).replace("{{ccID}}", id);
        body = body.replace("{{xPath}}", xpath);
        body = body.replace("{{text}}", node);
        body = body.replace("{{storeItemID}}", itemId);

        return body;
    }

    function getContentControlOoxml(body) {
        var ooxml = '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml" /></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p w:rsidR="00ED6193" w:rsidRDefault="006244BE">{{ooxml}}</w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>';
        return ooxml.replace("{{ooxml}}", body);
    }

    function addContentControlToDocument(ooxml, type) {
        Office.context.document.setSelectedDataAsync(
            ooxml,
            {
                coercionType: type == coercicionEnum.IMAGE
                    ? Office.CoercionType.Image
                    : Office.CoercionType.Ooxml
            },
            function (result) {
            }
        );
    }

    function createContentControl(node, xpath, type) {
        var ccID = generateContentControlUniqueId(node);
        var ccBodyTemplate = getContentControlBodyTemplate(type);
        var ccBody = getContentControlBody(ccBodyTemplate, ccID, node, xpath);
        var ooxml = getContentControlOoxml(ccBody);

        if (type == coercicionEnum.CONDITION) {
            localStorage.setItem("xpath", xpath);

            var dialog;
            Office.context.ui.displayDialogAsync(
                "https://localhost:44322/EditXPath.html",
                { height: 30, width: 50 },
                function (asyncResult) {
                    dialog = asyncResult.value;
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
                        dialog.close();

                        var newXPath = arg.message;
                        addContentControlToDocument(ooxml);
                        manageCustomXMLParts(type, ccID, newXPath);
                    });
                }
            );
        }
        else {
            addContentControlToDocument(ooxml, type);
            manageCustomXMLParts(type, ccID, xpath);
        }
    }

    function manageCustomXMLParts(ccType, ccID, xpath) {
        var xpathId = ccType == coercicionEnum.CONDITION
            ? generateContentControlUniqueId(ccID.split('_')[0])
            : ccID;

        updateXPathsCustomXmlPart(xpathsCustomXMLPartId, xpathId, xpath);

        if (ccType == coercicionEnum.CONDITION) {
            updateConditionsCustomXmlPart(conditionsCustomXMLPartId, ccID, xpathId);
        }
    }

    function updateXPathsCustomXmlPart(xmlPartId, xpathId, xpath) {
        Office.context.document.customXmlParts.getByIdAsync(xmlPartId, function (result) {
            var xmlPart = result.value;
            xmlPart.getNodesAsync('*', function (result) {
                var node = result.value[0];
                node.getXmlAsync(function (result) {
                    var xml = result.value;
                    xml = xml.replace(
                        "</xpaths>",
                        '<xpath id="' + xpathId + '"><dataBinding xpath="' + xpath + '" storeItemID="' + itemId + '"/></xpath></xpaths>'
                    );

                    node.setXmlAsync(xml, function (result) {
                    });
                });
            });
        });
    }

    function updateConditionsCustomXmlPart(xmlPartId, ccID, xpathRef) {
        Office.context.document.customXmlParts.getByIdAsync(xmlPartId, function (result) {
            var xmlPart = result.value;
            xmlPart.getNodesAsync('*', function (result) {
                var node = result.value[0];
                node.getXmlAsync(function (result) {
                    var xml = result.value;
                    xml = xml.replace(
                        "</conditions>",
                        '<condition id="' + ccID + '"><xpathref id="' + xpathRef + '"/></condition></conditions>'
                    );

                    node.setXmlAsync(xml, function (result) {
                        if (result.status !== "succeeded" || result.value.length === 0) {
                            displayError("Error when trying to set node XML");
                            return;
                        }
                    });
                });
            });
        });
    }

    function updateOldXPathWithNewXPart(xmlPartId, oldXPath, newXPath) {
        Office.context.document.customXmlParts.getByIdAsync(xmlPartId, function (result) {
            var xmlPart = result.value;
            xmlPart.getNodesAsync('*', function (result) {
                var node = result.value[0];
                node.getXmlAsync(function (result) {
                    var xml = result.value;
                    xml = xml.replace(oldXPath, newXPath)

                    node.setXmlAsync(xml, function (result) {
                    });
                });
            });
        });
    }

    function getXPathFromXMLByNodeId(_xml, nodeId) {
        var xml = $.parseXML(_xml);
        var xpaths = xml.childNodes[0].childNodes;
        for (var i = 0; i < xpaths.length; i++) {
            var xpath = xpaths[i];
            var xpathId = xpath.attributes[0].nodeValue;
            if (xpathId === nodeId) {
                var dataBinding = xpath.childNodes[0];
                var attributes = dataBinding.attributes;
                for (var j = 0; j < attributes.length; j++) {
                    var attribute = attributes[j];
                    if (attribute.nodeName === "xpath") {
                        return attribute.nodeValue;
                    }
                }
            }
        }
    }

    function getXPathRefFromXMLByNodeId(_xml, nodeId) {
        var xml = $.parseXML(_xml);
        var conditions = xml.childNodes[0].childNodes;
        for (var i = 0; i < conditions.length; i++) {
            var condition = conditions[i];
            var conditionId = condition.attributes[0].nodeValue;
            if (conditionId === nodeId) {
                var xPathRef = condition.childNodes[0];
                return xPathRef.attributes[0].nodeValue;
            }
        }
    }

    function getXPathByNodeId(namespace, nodeId) {
        if (namespace == "condition") {
            Office.context.document.customXmlParts.getByIdAsync(conditionsCustomXMLPartId, function (result) {
                var xmlPart = result.value;
                xmlPart.getNodesAsync('*', function (result) {
                    var node = result.value[0];
                    node.getXmlAsync(function (result) {
                        var xml = result.value;
                        var xPathRef = getXPathRefFromXMLByNodeId(xml, nodeId)

                        Office.context.document.customXmlParts.getByIdAsync(xpathsCustomXMLPartId, function (result) {
                            var xmlPart = result.value;
                            xmlPart.getNodesAsync('*', function (result) {
                                var node = result.value[0];
                                node.getXmlAsync(function (result) {
                                    var xml = result.value;
                                    var xPath = getXPathFromXMLByNodeId(xml, xPathRef);

                                    $("#display-xpath").text(xPath);
                                });
                            });
                        });
                    });
                });
            });
        }
        else {
            Office.context.document.customXmlParts.getByIdAsync(xpathsCustomXMLPartId, function (result) {
                var xmlPart = result.value;
                xmlPart.getNodesAsync('*', function (result) {
                    var node = result.value[0];
                    node.getXmlAsync(function (result) {
                        var xml = result.value;
                        var xPath = getXPathFromXMLByNodeId(xml, nodeId);

                        $("#display-xpath").text(xPath);
                    });
                });
            });
        }
    }

    function addOnClickEventHandler() {
        Office.context.document.addHandlerAsync(
            "documentSelectionChanged",
            function (e) {
                Word.run(function (context) {
                    var range = context.document.getSelection();
                    context.load(range);

                    return context.sync()
                        .then(function () {
                            var contentControl = range.parentContentControlOrNullObject;
                            context.load(contentControl, 'tag');

                            return context.sync()
                                .then(function () {
                                    if (!contentControl.m_isNull) {
                                        var tag = contentControl.tag;
                                        tag = tag.replace("od:", "").replace("&od:ContentType", "");

                                        var namespace = tag.split("=")[0];
                                        var nodeId = tag.split("=")[1];

                                        getXPathByNodeId(namespace, nodeId);
                                    }

                                    return context.sync();
                                });
                        });
                });
            },
            function (result) {

            }
        );
    }

    function addEditXPathOnClickEventHandler() {
        $("#edit-xpath").on("click", function () {
            var xpath = $("#display-xpath").text();
            if (xpath !== "undefined" && xpath !== "") {
                localStorage.setItem("xpath", xpath);

                var dialog;
                Office.context.ui.displayDialogAsync(
                    "https://localhost:44322/EditXPath.html",
                    { height: 30, width: 50 },
                    function (asyncResult) {
                        dialog = asyncResult.value;
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
                            dialog.close();

                            var newXPath = arg.message;
                            $("#display-xpath").text(newXPath);
                            updateOldXPathWithNewXPart(xpathsCustomXMLPartId, xpath, newXPath);
                        });
                    }
                );
            }
        });
    }

    function testTemplate() {
        debugger;

        var xhr = new XMLHttpRequest();
        xhr.withCredentials = true;

        xhr.onload = function () {
            var test = xhr.responseText;
            debugger;
        };

        xhr.open("POST", "http://svnsidvjv07:11300/totem/api/totem-client/test-template/2086");
        xhr.setRequestHeader("Authorization", "Bearer eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiJhZG1pbiIsImV4cCI6MTU5NjQ1OTAzMywiYXV0aCI6IkdST1VQX0FETUlOLFJPTEVfQURNSU4ifQ.QTiB59vWpq8fmF5EGBZMwn9IGJVoB74Hw0lVVi2oyrrSt4kIVy8u5Ddzzu1geiKM9GqztTS5HhV-t8pojqYbEw");
        xhr.setRequestHeader("Accept", "*/*");
        xhr.setRequestHeader("Accept-Language", "fr-BE,fr-FR;q=0.9,fr;q=0.8,en-US;q=0.7,en;q=0.6");

        xhr.send();
    }

    $(function () {
        //checkIfCustomXMLPartsExist();
        //generateDataDefinitionTree();
        //addOnClickEventHandler();
        //addEditXPathOnClickEventHandler();

        testTemplate();
    });

};





function testTemplate() {
    var xhr = new XMLHttpRequest();
    xhr.withCredentials = true;

    xhr.onload = function () {
        var test = xhr.responseText;
        debugger;
    };

    xhr.open("POST", "http://svnsidvjv07:11300/totem/api/totem-client/test-template/2086");
    xhr.setRequestHeader("Authorization", "Bearer eyJhbGciOiJIUzUxMiJ9.eyJzdWIiOiJhZG1pbiIsImV4cCI6MTU5NjQ1OTAzMywiYXV0aCI6IkdST1VQX0FETUlOLFJPTEVfQURNSU4ifQ.QTiB59vWpq8fmF5EGBZMwn9IGJVoB74Hw0lVVi2oyrrSt4kIVy8u5Ddzzu1geiKM9GqztTS5HhV-t8pojqYbEw");
    xhr.setRequestHeader("Accept", "*/*");
    xhr.setRequestHeader("Accept-Language", "fr-BE,fr-FR;q=0.9,fr;q=0.8,en-US;q=0.7,en;q=0.6");

    xhr.send();
}

$(function () {
    //checkIfCustomXMLPartsExist();

    //generateDataDefinitionTree();
    //addOnClickEventHandler();
    //addEditXPathOnClickEventHandler();

    debugger;

    testTemplate();
});
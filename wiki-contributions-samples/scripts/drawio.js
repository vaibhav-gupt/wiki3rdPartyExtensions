VSS.init({ explicitNotifyLoaded: true, usePlatformScripts: true, usePlatformStyles: true, extensionReusedCallback: registerContribution });

// We need to register the new contribution if this extension host is reused
function registerContribution(contribution) {
    // Register the fully-qualified contribution id here.
    // Because we're using the contribution id, we do NOT need to define a registeredObjectId in the extension manfiest.
    VSS.register(contribution.id, {
        sendContent: function (content, isEditMode) {
            updateImageSrc(content);
            resizeSelf();
            document.removeEventListener('dblclick', registerEditor);
            if (isEditMode) {
                document.addEventListener('dblclick', registerEditor);
            }
        }
    });
}

var registerEditor = function (evt) {
    // Edits an image with drawio class on double click
    var source = evt.srcElement || evt.target;
    if (source.nodeName == 'IMG' && source.className == 'drawio') {
        if (source.drawIoWindow == null || source.drawIoWindow.closed) {
            // Implements protocol for loading and exporting with embedded XML
            var receive = function (evt) {
                if (evt.data.length > 0 && evt.source == source.drawIoWindow) {
                    var msg = JSON.parse(evt.data);

                    // Received if the editor is ready
                    if (msg.event == 'init') {
                        // Sends the data URI with embedded XML to editor
                        source.drawIoWindow.postMessage(
                            JSON.stringify({ action: 'load', xmlpng: source.getAttribute('src') }), '*');
                    }
                    // Received if the user clicks save
                    else if (msg.event == 'save') {
                        // Sends a request to export the diagram as XML with embedded PNG
                        source.drawIoWindow.postMessage(JSON.stringify(
                            { action: 'export', format: 'xmlpng', spinKey: 'saving' }), '*');
                    }
                    // Received if the export request was processed
                    else if (msg.event == 'export') {
                        updateImageSrc(msg.data);
                        VSS.getService("ms.vss-wiki-web.wiki-extension-data-service").then(function (obj) {
                            obj.setData({ elementId: VSS.getConfiguration(), graphData: msg.data });
                            console.log("data set message has been sent");
                        });
                    }

                    // Received if the user clicks exit or after export
                    if (msg.event == 'exit' || msg.event == 'export') {
                        VSS.getService("ms.vss-wiki-web.wiki-extension-data-service").then(function (obj) {
                            obj.exitFullScreen({ elementId: VSS.getConfiguration() }).then(function (completed) {
                                if (completed) {
                                    resizeSelf();
                                }
                            });
                            console.log("exitFullScreen message has been sent");
                        });

                        // Closes the editor
                        window.removeEventListener('message', receive);
                        source.drawIoWindow.close();
                        source.drawIoWindow = null;
                        document.body.removeChild(iframe);
                    }
                }
            };

            window.addEventListener('message', receive);
            // Opens the editor
            var iframe = document.createElement('iframe');
            iframe.setAttribute('frameborder', '0');
            iframe.setAttribute('id', 'editor');
            iframe.setAttribute('src', 'https://www.draw.io/?embed=1&ui=atlas&spin=1&modified=unsavedChanges&proto=json');
            document.body.appendChild(iframe);
            source.drawIoWindow = iframe.contentWindow;
            resizeSelf(true);
            VSS.getService("ms.vss-wiki-web.wiki-extension-data-service").then(function (obj) {
                obj.setFullScreen({ elementId: VSS.getConfiguration() });
            });
        }
        else {
            // Shows existing editor window
            source.drawIoWindow.focus();
        }
    }
}

var resizeSelf = function (isEditMode) {
    if(isEditMode) {
      VSS.resize(-1, -1);
    }
    else {
      var imageElement = document.getElementsByClassName("drawio")[0];
      window.setTimeout(function(){
        VSS.resize(imageElement.offsetWidth + 20, imageElement.offsetHeight + 20);
      }, 0);
    }
}

var updateImageSrc = function (content) {
    var imageElement = document.getElementsByClassName("drawio")[0];
    // Updates the data URI of the image
    imageElement.setAttribute('src', content);
}

// Show context info when ready
VSS.ready(function () {
    registerContribution(VSS.getContribution());
    VSS.notifyLoadSucceeded();
});
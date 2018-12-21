VSS.init({explicitNotifyLoaded: true, usePlatformScripts: true, usePlatformStyles: true, extensionReusedCallback: registerContribution });
 
 // Edits an image with drawio class on double click
 document.addEventListener('dblclick', function(evt)
 {
   var url = 'https://www.draw.io/?embed=1&ui=atlas&spin=1&modified=unsavedChanges&proto=json';
   var source = evt.srcElement || evt.target;

   if (source.nodeName == 'IMG' && source.className == 'drawio')
   {
     if (source.drawIoWindow == null || source.drawIoWindow.closed)
     {
        var iframe = document.createElement('iframe');
        iframe.setAttribute('frameborder', '0');
        iframe.setAttribute('id', 'editor');

       // Implements protocol for loading and exporting with embedded XML
       var receive = function(evt)
       {
         if (evt.data.length > 0 && evt.source == source.drawIoWindow)
         {
           var msg = JSON.parse(evt.data);

           // Received if the editor is ready
           if (msg.event == 'init')
           {
             // Sends the data URI with embedded XML to editor
             source.drawIoWindow.postMessage(
                 JSON.stringify({action: 'load', xmlpng: source.getAttribute('src')}), '*');
           }
           // Received if the user clicks save
           else if (msg.event == 'save')
           {
             // Sends a request to export the diagram as XML with embedded PNG
             source.drawIoWindow.postMessage(JSON.stringify(
               {action: 'export', format: 'xmlpng', spinKey: 'saving'}), '*');
           }
           // Received if the export request was processed
           else if (msg.event == 'export')
           {
             // Updates the data URI of the image
             source.setAttribute('src', msg.data);
             VSS.getService("ms.vss-wiki-web.wiki-extension-data-service").then(function(obj) {
                obj.setData({elementId: VSS.getConfiguration(), graphData: msg.data});
                console.log("data set message has been sent");
             });
           }
                        
           // Received if the user clicks exit or after export
           if (msg.event == 'exit' || msg.event == 'export')
           {
             // Closes the editor
             window.removeEventListener('message', receive);
             source.drawIoWindow.close();
             source.drawIoWindow = null;
             document.body.removeChild(iframe);
           }
         }
       };

       window.addEventListener('message', receive);
       VSS.getService("ms.vss-wiki-web.wiki-extension-data-service").then(function(obj) {
            obj.setFullScreen({elementId: VSS.getConfiguration()});
        });
       // Opens the editor
       iframe.setAttribute('src', url);
       document.body.appendChild(iframe);
       source.drawIoWindow = iframe.contentWindow;
     }
     else
     {
       // Shows existing editor window
       source.drawIoWindow.focus();
     }
   }
 });
 
var close = function()
{
    window.removeEventListener('message', receive);
    document.body.removeChild(iframe);
}; 

// We need to register the new contribution if this extension host is reused
function registerContribution(contribution) {
    // Register the fully-qualified contribution id here.
    // Because we're using the contribution id, we do NOT need to define a registeredObjectId in the extension manfiest.
    VSS.register(contribution.id, {
        sendContent: function(content){console.log(content);}
    });
}

// Show context info when ready
VSS.ready(function () {
    registerContribution(VSS.getContribution());
    VSS.notifyLoadSucceeded();
});

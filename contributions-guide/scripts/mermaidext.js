VSS.init({explicitNotifyLoaded: true, usePlatformScripts: true, usePlatformStyles: true, extensionReusedCallback: registerContribution });

mermaid.initialize({
    startOnLoad:false
});

// We need to register the new contribution if this extension host is reused
function registerContribution(contribution) {
        // Register the fully-qualified contribution id here.
        // Because we're using the contribution id, we do NOT need to define a registeredObjectId in the extension manfiest.
        VSS.register(contribution.id, {
            sendContent: sendContent
        });
}

function sendContent(content) {
    var insertSvg = function(svgCode, bindFunctions){
        element;
        document.getElementById("merm1").innerHTML = svgCode;
    };

    var graphDefinition = 'graph TB\na-->b';
    mermaid.render('graphDiv', content, insertSvg);
    VSS.resize();
}

// Show context info when ready
VSS.ready(function () {
    registerContribution(VSS.getContribution());
    VSS.notifyLoadSucceeded();
});


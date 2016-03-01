function spoHelpers() {
    var setMasterPage = function(masterPageUrl) {
        var clientContext = SP.ClientContext.get_current();
        var website = clientContext.get_web();

        clientContext.load(website);
        clientContext.executeQueryAsync(function() { 
            //do nothing
        }, function() {console.log("fail")});
                
        // var masterPageUrl = "/_catalogs/masterpage/oslo.master";
        
        website.set_masterUrl(masterPageUrl);
        website.update();
        clientContext.load(website);
        
        clientContext.executeQueryAsync(function() { 
            console.log("New MasterUrl for " + website.get_title() +": " + website.get_masterUrl())
            
            window.website = website;
        }, function(sender, args) {console.log("fail:"); console.log(args.get_message())});
    }

    var setAlternateCssUrl = function(alternateCssUrl) {
        var clientContext = SP.ClientContext.get_current();
        var website = clientContext.get_web();

        clientContext.load(website);
        clientContext.executeQueryAsync(function() {
            //do nothing
        }, function() {console.log("fail")});
                
        //var alternateCssUrl = "/_catalogs/masterpage/oslo.master";
        
        website.set_alternateCssUrl(alternateCssUrl);
        website.update();
        clientContext.load(website);
        
        clientContext.executeQueryAsync(function() { 
            console.log("New AlternateCssUrl for " + website.get_title() +": " + website.get_alternateCssUrl())
            
            window.website = website;
        }, function(sender, args) {console.log("fail:"); console.log(args.get_message())});
    }
    
    return {
        setAlternateCssUrl : setAlternateCssUrl,
        setMasterPage : setMasterPage
    }
}

var onload = function() { 
    
    window.spoHelpers = spoHelpers();
    
	console.log("SharePoint Online Debug Window Helpers Loaded");
};

// SP.SOD.registerSod("jquery", "https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js");
// SP.SOD.executeFunc("jquery", null, onload);
// SP.SOD.executeOrDelayUntilScriptLoaded(onload, "sp.js");

SP.SOD.executeFunc('sp.js', 'SP.ClientContext', onload);

 
var onload = function() { 
	console.log("Loading Data..");
	
	//var clientContext = new SP.ClientContext("https://demro.sharepoint.com");
	var clientContext = SP.ClientContext.get_current();
	var website = clientContext.get_web();

	clientContext.load(website);
	clientContext.executeQueryAsync(function() { 
		console.log(website.get_title())
		
		window.website = website;
	}, function() {console.log("fail")});
			
	
	var masterPageUrl = "/_catalogs/masterpage/oslo.master";
	
	website.set_masterUrl(masterPageUrl);
	website.update();
	clientContext.load(website);
	
	clientContext.executeQueryAsync(function() { 
		console.log(website.get_masterUrl())
		
		window.website = website;
	}, function(sender, args) {console.log("fail:"); console.log(args.get_message())});
	
	
};

// SP.SOD.registerSod("jquery", "https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js");
// SP.SOD.executeFunc("jquery", null, onload);
// SP.SOD.executeOrDelayUntilScriptLoaded(onload, "sp.js");

SP.SOD.executeFunc('sp.js', 'SP.ClientContext', onload);

var SP = SP || {};

var onload = function() { 
    
    window.spoHelpers = new SpoHelpers();
    
	console.log("SharePoint Online Debug Window Helpers Loaded");
};

// SP.SOD.registerSod("jquery", "https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js");
// SP.SOD.executeFunc("jquery", null, onload);
// SP.SOD.executeOrDelayUntilScriptLoaded(onload, "sp.js");

SP.SOD.executeFunc('sp.js', 'SP.ClientContext', onload);




/* Spo Helpers
 * This class is a collection of static methods to help QoL when prototyping in O365 through the debug window
 */
function SpoHelpers() {
    
    var _ctx = {};
    var _web = {};
    this.web = {}
    this.ctx = {}
    
    var init = function(obj) { 
        _ctx = SP.ClientContext.get_current();
        _web = _ctx.get_web();
        _ctx.load(_web);
        _ctx.executeQueryAsync(function() { 
            obj.web = _web;
            obj.ctx = _ctx;
            console.log("Initialized spoHelpers")
        }, function() {console.log("fail")});
    }
    init(this);
    
    this.initialize = function() {
        init();
    };    

    
    this.setMasterPage = function(masterPageUrl) {        
        _web.set_masterUrl(masterPageUrl);
        _web.update();
        _ctx.load(_web);
        
        _ctx.executeQueryAsync(function() { 
            
            console.log("New MasterUrl for " + _web.get_title() +": " + _web.get_masterUrl())
            
        }, function(sender, args) {console.log("fail:"); console.log(args.get_message())});
    }

    this.setAlternateCssUrl = function(alternateCssUrl) {
        _web.set_alternateCssUrl(alternateCssUrl);
        _web.update();
        _ctx.load(_web);
        
        _ctx.executeQueryAsync(function() { 
            
            console.log("New AlternateCssUrl for " + _web.get_title() +": " + _web.get_alternateCssUrl())
            
        }, function(sender, args) {console.log("fail:"); console.log(args.get_message())});
    }
    
    // return {
    //     setAlternateCssUrl : setAlternateCssUrl,
    //     setMasterPage : setMasterPage,
    //     setJavascriptViaCustomAction : setJavascriptViaCustomAction
    // }
}
 
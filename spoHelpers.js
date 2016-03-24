var SP = SP || {};

var onload = function() { 
    
    // First, checks if it isn't implemented yet.
    if (!String.prototype.format) {
    String.prototype.format = function() {
        var args = arguments;
        return this.replace(/{(\d+)}/g, function(match, number) { 
        return typeof args[number] != 'undefined'
            ? args[number]
            : match
        ;
        });
    };
    }
    
    window.spo = new SpoHelpers();
    
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
    
    this.load = function(obj) {
        _ctx.load(obj);
        _ctx.executeQueryAsync(function() { console.log("loaded")}, function(sender, args) {console.log('Request failed. ' + args.get_message() + 
        '\n' + args.get_stackTrace());})
        
        return;
    }
    
    // this.executeQuerySyncronous = function() {  
    //     var dfd = new jQuery.Deferred();
        
    //     _ctx.executeQueryAsync(function() { console.log("complete"); dfd.resolve()}, function(sender, args) {console.log('Request failed. ' + args.get_message() + 
    //     '\n' + args.get_stackTrace()); dfd.resolve();})
        
    //     var count = 0;
    //     while(dfd.state() == "pending") {
    //         if(count > 10) {
    //             break;
    //         }
    //         setTimeout(function() {console.log("processing..")}, 500);
    //     }
        
    //     return; 
    // }
    
    var deleteWebPart = function (wpm, id) {
        var dfd = jQuery.Deferred();
        
        if(id == "" || id == undefined) {
            console.log("webpart id is null, nothing to delete");
            dfd.resolve();;
            return dfd.promise();
        }
        
        var webParts = wpm.get_webParts();
        var wpToDelete = webParts.getById(id);
        wpToDelete.deleteWebPart();
        
        spo.ctx.load(wpToDelete);
        spo.ctx.executeQueryAsync(function() {
            console.log("deleted webpart");
            dfd.resolve();
        });
        
        return dfd.promise();
    }
    
    var getWebPartFromWebPartManager = function(wpm, webPartTitle) {
            var dfd = new jQuery.Deferred();
            var existingWebParts = wpm.get_webParts();
            spo.ctx.load(existingWebParts);
            spo.ctx.executeQueryAsync(function() {
            
                var data = existingWebParts.get_data()
                var realWebParts = [];
                var wpId = "";
                data.forEach(function(webPart) {
                    var realWebPart = webPart.get_webPart() 
                    spo.ctx.load(realWebPart);
                    realWebParts.push({"webpart":realWebPart, "id":webPart.get_id()});
                }, this); 
                spo.ctx.executeQueryAsync(function() {
                
                    realWebParts.forEach(function(wp) {
                        if(wp.webpart.get_title() == webPartTitle) {
                            console.log("found web part:" + webPartTitle + " " + wp.id); 
                            wpId = wp.id;
                        }
                    }, this);
                    
                    window.webPartId = wpId;
                    dfd.resolve(wpId);
                });
            });
            
            return dfd.promise();
    }
    
    
    var wpXml = '<webParts> '
                + '   <webPart xmlns="http://schemas.microsoft.com/WebPart/v3"> '
                + '     <metaData> '
                + '       <type name="Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" /> '
                + '       <importErrorMessage>Cannot import this Web Part.</importErrorMessage> '
                + '     </metaData> '
                + '     <data> '
                + '       <properties> '
                + '         <property name="ExportMode" type="exportmode">All</property> '
                + '         <property name="HelpUrl" type="string" /> '
                + '         <property name="Hidden" type="bool">False</property> '
                + '         <property name="Description" type="string">Allows authors to insert HTML snippets or scripts.</property> '
                + '         <property name="Content" type="string">&lt;div id="embedded-my-feed" style="height:400px;width:500px;"&gt;&lt;/div&gt;  '
                + '     &lt;script type="text/javascript" src="https://c64.assets-yammer.com/assets/platform_embed.js"&gt;&lt;/script&gt; '
                + '     &lt;script \'type="text/javascript"&gt; yam.connect.embedFeed({   '
                + '               container: \'#embedded-my-feed\', '
                + '               network: \'morganfranklin.com\'  }); '
                + '     &lt;/script&gt;</property> '
                + '         <property name="CatalogIconImageUrl" type="string" /> '
                + '         <property name="Title" type="string">Yammer Feed</property> '
                + '         <property name="AllowHide" type="bool">True</property> '
                + '         <property name="AllowMinimize" type="bool">True</property> '
                + '         <property name="AllowZoneChange" type="bool">True</property> '
                + '         <property name="TitleUrl" type="string" /> '
                + '         <property name="ChromeType" type="chrometype">TitleOnly</property> '
                + '         <property name="AllowConnect" type="bool">True</property> '
                + '         <property name="Width" type="unit" /> '
                + '         <property name="Height" type="unit" /> '
                + '         <property name="HelpMode" type="helpmode">Navigate</property> '
                + '         <property name="AllowEdit" type="bool">True</property> '
                + '         <property name="TitleIconImageUrl" type="string" /> '
                + '         <property name="Direction" type="direction">NotSet</property> '
                + '         <property name="AllowClose" type="bool">True</property> '
                + '         <property name="ChromeState" type="chromestate">Normal</property> '
                + '       </properties> '
                + '     </data> '
                + '   </webPart> '
                + ' </webParts> ';
    
    var addWebPartToWebPartManager = function(wpm, wpXml) {
        var dfd = new jQuery.Deferred();
        var importedWebPart = wpm.importWebPart(wpXml);
        var webPartAdded = wpm.addWebPart(importedWebPart.get_webPart(), "wpz", 0);
        
        spo.ctx.load(webPartAdded)
        spo.ctx.executeQueryAsync(function() {
           console.log("loaded webpart to webpart manager");
           dfd.resolve(webPartAdded); 
        });
        
        return dfd.promise();
    }
    
    var replaceWebPartWikiPage = function(wpm, page, webPartAddedId) {
        var dfd = new jQuery.Deferred();
        var marker = "<div id='yammerWidget' class=\"ms-rtestate-read ms-rte-wpbox\" contentEditable=\"false\"><div class=\"ms-rtestate-read {0}\" id=\"div_{0}\"></div><div style='display:none' id=\"vid_{0}\"></div></div>".format(webPartAddedId);
        var pageListItem = page.get_listItemAllFields()
        
        spo.ctx.load(pageListItem);
        spo.ctx.executeQueryAsync(function() {
            console.log("loaded page list item");
        
            var wikiField = pageListItem.get_fieldValues().WikiField;
            
            //TODO remove this business logic from this method
            var pageElem = jQuery(wikiField);
            
            var yammerParent = "";
            
            if(pageElem.find("#div_840b70a6-63af-4e80-8c55-53a7a5b3a858").length > 0) {
                yammerParent = pageElem.find("#div_840b70a6-63af-4e80-8c55-53a7a5b3a858").parent();
            } else if(pageElem.find("#div_5960e2d0-6996-41a8-a6d5-b2caa9c24e12").length > 0) {
                yammerParent = pageElem.find("#div_5960e2d0-6996-41a8-a6d5-b2caa9c24e12").parent();
            } else {
                yammerParent = pageElem.find("#yammerWidget");
            }
            
            yammerParent.replaceWith(marker);
            
            pageListItem.set_item('WikiField', pageElem.html()); 
            pageListItem.update();
            spo.ctx.executeQueryAsync(function() {
                dfd.resolve();    
            });
        });
    
        return dfd.promise();
    }
    
    this.console = {log: function(message) {
        if(true) {
            window.console.log(message);
        }
    }}
    
    this.test = function(webPartTitle) {
        webPartTitle = "Yammer Feed";
        var page = spo.web.getFileByServerRelativeUrl(_spPageContextInfo.serverRequestPath);
        //var page = spo.web.getFileByServerRelativeUrl("/sites/acctmgmt/DemroTest-DemroTestProjectII/SitePages/Home_copy(5).aspx");
        var wpm = page.getLimitedWebPartManager(1);
        spo.ctx.load(wpm);
        
        
        
        console.log("Loading web part manager");
        spo.ctx.executeQueryAsync(function() {
            console.log("Getting existing webpart");
            getWebPartFromWebPartManager(wpm, webPartTitle).done(function(wpId) {
                console.log("deleting existing webpart");
                deleteWebPart(wpm, wpId).done(function() {
                    console.log("adding new webpart to wpm");
                    addWebPartToWebPartManager(wpm, wpXml).done(function(webPartAdded) {
                        console.log("updating wiki page UI with new webpart");
                        replaceWebPartWikiPage(wpm, page, webPartAdded.get_id()).done(function() {
                        //replaceWebPartWikiPage(wpm, page, "12345").done(function() {
                            console.log("Complete!");
                        });    
                    });
                     
                    
                });
            });
        });
    }
    
    // this.addWebPart = function() {
    //     //get current page
        
    //     //var page = spo.web.getFileByServerRelativeUrl("/sites/acctmgmt/DemroTest-DemroTest/SitePages/Home_copy.aspx");
    //     var page = spo.web.getFileByServerRelativeUrl("/sites/acctmgmt/DemroTest-DemroTestProjectII/SitePages/Home_copy(5).aspx");
    //     //var page = spo.web.getFileByServerRelativeUrl("/sites/demroTeam/test2/sitepages/home.aspx");
    //     var wpm = page.getLimitedWebPartManager(1);
    //     var webParts = wpm.get_webParts();
        
    //     spo.ctx.load(page);
    //     spo.ctx.load(wpm);
    //     spo.ctx.load(webParts);
        
    //     spo.ctx.executeQueryAsync(function() {
    //         console.log("loaded page");
            
    //         /******* */
            
    //         var wpXml = '<webParts> '
    //             + '   <webPart xmlns="http://schemas.microsoft.com/WebPart/v3"> '
    //             + '     <metaData> '
    //             + '       <type name="Microsoft.SharePoint.WebPartPages.ScriptEditorWebPart, Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" /> '
    //             + '       <importErrorMessage>Cannot import this Web Part.</importErrorMessage> '
    //             + '     </metaData> '
    //             + '     <data> '
    //             + '       <properties> '
    //             + '         <property name="ExportMode" type="exportmode">All</property> '
    //             + '         <property name="HelpUrl" type="string" /> '
    //             + '         <property name="Hidden" type="bool">False</property> '
    //             + '         <property name="Description" type="string">Allows authors to insert HTML snippets or scripts.</property> '
    //             + '         <property name="Content" type="string">&lt;div id="embedded-my-feed" style="height:400px;width:500px;"&gt;&lt;/div&gt;  '
    //             + '     &lt;script type="text/javascript" src="https://c64.assets-yammer.com/assets/platform_embed.js"&gt;&lt;/script&gt; '
    //             + '     &lt;script \'type="text/javascript"&gt; yam.connect.embedFeed({   '
    //             + '               container: \'#embedded-my-feed\', '
    //             + '               network: \'morganfranklin.com\'  }); '
    //             + '     &lt;/script&gt;</property> '
    //             + '         <property name="CatalogIconImageUrl" type="string" /> '
    //             + '         <property name="Title" type="string">Yammer Feed</property> '
    //             + '         <property name="AllowHide" type="bool">True</property> '
    //             + '         <property name="AllowMinimize" type="bool">True</property> '
    //             + '         <property name="AllowZoneChange" type="bool">True</property> '
    //             + '         <property name="TitleUrl" type="string" /> '
    //             + '         <property name="ChromeType" type="chrometype">TitleOnly</property> '
    //             + '         <property name="AllowConnect" type="bool">True</property> '
    //             + '         <property name="Width" type="unit" /> '
    //             + '         <property name="Height" type="unit" /> '
    //             + '         <property name="HelpMode" type="helpmode">Navigate</property> '
    //             + '         <property name="AllowEdit" type="bool">True</property> '
    //             + '         <property name="TitleIconImageUrl" type="string" /> '
    //             + '         <property name="Direction" type="direction">NotSet</property> '
    //             + '         <property name="AllowClose" type="bool">True</property> '
    //             + '         <property name="ChromeState" type="chromestate">Normal</property> '
    //             + '       </properties> '
    //             + '     </data> '
    //             + '   </webPart> '
    //             + ' </webParts> ';
            
            
    //         var importedWebPart = wpm.importWebPart(wpXml);
            
            
            

            
    //         var webPartAdded = wpm.addWebPart(importedWebPart.get_webPart(), "wpz", 0);
    //         spo.ctx.load(webPartAdded)
    //         spo.ctx.executeQueryAsync(function() {
    //             console.log("loaded webpart")
            
            
    //             var webPartAddedId = webPartAdded.get_id()
    //             window.webPartAdded = webPartAdded;
    //             var marker = "<div id='yammerWidget' class=\"ms-rtestate-read ms-rte-wpbox\" contentEditable=\"false\"><div class=\"ms-rtestate-read {0}\" id=\"div_{0}\"></div><div style='display:none' id=\"vid_{0}\"></div></div>".format(webPartAddedId);
    //             var item = page.get_listItemAllFields()
                
                
    //             spo.ctx.load(item);
    //             spo.ctx.executeQueryAsync(function() {
    //                 console.log("loaded page list item");
                    
    //                 var wikiField = item.get_fieldValues().WikiField;
                    
                    
    //                 var page = jQuery(wikiField);
    //                 var yammer = page.find("#div_840b70a6-63af-4e80-8c55-53a7a5b3a858");
    //                 //var yammer = page.find("#yammer");
    //                 var yammerParent = yammer.parent();
    //                 yammerParent.replaceWith(marker);
                    
                    
    //                 //var newField = marker + wikiField;
                    
    //                 window.wikiField = wikiField;
    //                 window.marker = marker;
                    
    //                 item.set_item('WikiField', page.html()); 
    //                 item.update();
    //                 spo.ctx.executeQueryAsync();
                            
    //             });
            
    //         });
    //     });
        
        
        
        
    //} 
    
    // return {
    //     setAlternateCssUrl : setAlternateCssUrl,
    //     setMasterPage : setMasterPage,
    //     setJavascriptViaCustomAction : setJavascriptViaCustomAction
    // }
}
 
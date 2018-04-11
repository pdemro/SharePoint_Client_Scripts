var SP = SP || {};

var onload = function() { 
    
    // Add string.format to page
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

SP.SOD.registerSod("jquery", "https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js");

SP.SOD.executeFunc("jquery", '$.attr', function() {
    SP.SOD.executeFunc('sp.js', 'SP.ClientContext', onload);
});



/* Spo Helpers
 * This class is a collection of static methods to help QoL when prototyping in O365 through the debug window
 * TODO:  Break this out into multiple helper classes based on functionality
 * TODO: Convert me to typescript
 */
function SpoHelpers() {
    
    var _this = this;
    var _ctx = {};
    var _web = {};
    var _site = {};
    this.web = {};
    this.site = {};
    this.ctx = {};

    /** Constructor **/
    var init = function(obj) { 
        var dfd = new jQuery.Deferred();
        _ctx = SP.ClientContext.get_current();
        _web = _ctx.get_web();
        _site = _ctx.get_site();
        _ctx.load(_web);
        _ctx.executeQueryAsync(function() { 
            obj.web = _web;
            obj.ctx = _ctx;
            obj.site = _site;
            console.log("Initialized spoHelpers")
            dfd.resolve();
        }, function() 
            {
                console.log("fail")
                dfd.resolve();
            }
        );
        return dfd.promise();
    }
    init(this);
    
    this.initialize = function() {
        return init(_this);
    };    
    

    /** File Publishing **/
    function publishFile(file) {
        // var postUrl = `${serverUrl}/_api/web/GetFileByServerRelativeUrl('${file.d.ServerRelativeUrl}')/checkin(comment='Check-in comment for the publish operation.',checkintype=0)`;
         var postUrl = _spPageContextInfo.webAbsoluteUrl + "/_api/web/GetFileByServerRelativeUrl('" + file.get_serverRelativeUrl() + "')/publish(comment='Check-in comment for the publish operation.')";
         var dfd = jQuery.Deferred();

        jQuery.ajax({
            url: postUrl,
            type: "POST",
            headers: {
                "X-RequestDigest": jQuery("#__REQUESTDIGEST").val(),
                "content-type": "application/json;odata=verbose",
                //"content-length": body.length,
                //"IF-MATCH": itemMetadata.etag,
                "X-HTTP-Method": "MERGE"
            }
        }).always(function() {
            dfd.resolve();
        });

        return dfd.promise();
    }

    function checkInFile(file) {
        // var postUrl = `${serverUrl}/_api/web/GetFileByServerRelativeUrl('${file.d.ServerRelativeUrl}')/checkin(comment='Check-in comment for the publish operation.',checkintype=0)`;
         var postUrl = _spPageContextInfo.webAbsoluteUrl + "/_api/web/GetFileByServerRelativeUrl('" + file.get_serverRelativeUrl() + "')/checkin(comment='Check-in comment for the publish operation.', checkintype=0)";

         var dfd = jQuery.Deferred();

        jQuery.ajax({
            url: postUrl,
            type: "POST",
            headers: {
                "X-RequestDigest": jQuery("#__REQUESTDIGEST").val(),
                "content-type": "application/json;odata=verbose",
                //"content-length": body.length,
                //"IF-MATCH": itemMetadata.etag,
                "X-HTTP-Method": "MERGE"
            }
        }).always(function() {
            dfd.resolve(file);
        });

        return dfd.promise();
    }

    function recursivePublishFiles(folder) {
        //console.log(folder.get_name());
        var subFolders = folder.get_folders();
        spo.load(subFolders).done(function() {
            if(subFolders.get_count() > 0) {
                subFolders.get_data().forEach(function(v, i) {
                    //console.log("subfolder: " + v.get_name()); 
                    recursivePublishFiles(v);
                })
            }
        });

        var files = folder.get_files();
        spo.load(files).done(function() {
            files.get_data().forEach(function(v, i) {
                //console.log("file: " + v.get_name());
                if(v.get_minorVersion() > 0) {
                    console.log("file: " + v.get_name());
                    checkInFile(v).done(function(checkedInFile) {
                        publishFile(v);
                    })
                }
                
            });
        });
    }

    this.publishAllFiles = function(folderUrl) {
        var folder = spo.web.getFolderByServerRelativeUrl(folderUrl);

        this.load(folder).done(function() {
            recursivePublishFiles(folder);
        })
    }


    /** General Tools **/

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

    this.getCustomActionBySequence = function (ctxObj, sequence) {
        var dfd = jQuery.Deferred();
        var customActions = ctxObj.get_userCustomActions();

        this.load(customActions).done(function() {
            var customAction = null;
            var data = customActions.get_data();
            for(var i = 0; i < customActions.get_count(); i++) { 
                var action = data[i];

                if(action.get_sequence() == sequence) {
                    customAction = action;
                }                                
            }
            dfd.resolve(customAction);
        });

        return dfd.promise();
    }

    this.deleteCustomActionBySequence = function(ctxObj, sequence) {
        var dfd = jQuery.Deferred();
        
        this.getCustomActionBySequence(ctxObj, sequence).done(function(customAction) {
            
            if(customAction) {
                customAction.deleteObject();
                _ctx.load(customAction);
                _ctx.executeQueryAsync(function() {
                    console.log("deleted custom action");
                })
            } else {
                console.log("CustomAction with sequence {0} does not exist".format(sequence));
            }

            dfd.resolve();
        })

        return dfd.promise();
    }

    this.addCustomActionScriptBlock = function(ctxObj, title, sequence, scriptBlock) {

        this.deleteCustomActionBySequence(ctxObj, sequence).done(function(customAction) {
            var customActions = ctxObj.get_userCustomActions();
            customAction = customActions.add();
            
            customAction.set_location("ScriptLink");
            customAction.set_scriptBlock(scriptBlock);
            customAction.set_title(title);
            customAction.set_sequence(sequence);

            customAction.update();
            _ctx.executeQueryAsync();
        });
    }
    
    this.addCustomActionSpFx = function(ctxObj, location, title, sequence, componentId, propertiesJson) {

        this.deleteCustomActionBySequence(ctxObj, sequence).done(function(customAction) {
            var customActions = ctxObj.get_userCustomActions();
            customAction = customActions.add();
            
            customAction.set_location(location);
            customAction.set_title("Global Nav");
            customAction.set_sequence(sequence);
	    customAction.set_clientSideComponentId(componentId);
	    customAction.set_clientSideComponentProperties(propertiesJson);

            customAction.update();
            _ctx.executeQueryAsync();
        });
    }
    
    this.load = function(obj) {
        var dfd = new jQuery.Deferred();
        _ctx.load(obj);
        _ctx.executeQueryAsync(function() { console.log("loaded"); dfd.resolve();}, function(sender, args) {console.log('Request failed. ' + args.get_message() + 
        '\n' + args.get_stackTrace()); dfd.resolve();})
        
        
        return dfd.promise();
    }
    
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
                + '               network: \'company.com\'  }); '
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
                console.log("project style");
                yammerParent = pageElem.find("#div_840b70a6-63af-4e80-8c55-53a7a5b3a858").parent();
            } else if(pageElem.find("#div_5960e2d0-6996-41a8-a6d5-b2caa9c24e12").length > 0) {
                console.log("client style")
                yammerParent = pageElem.find("#div_5960e2d0-6996-41a8-a6d5-b2caa9c24e12").parent();
            } else {
                console.log("new style");
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
    
    this.replaceWebpartOnPage = function(webPartTitle, siteAbsoluteUrl, pageRelativeUrl) {
        //webPartTitle = "Yammer Feed";
        //var page = spo.web.getFileByServerRelativeUrl(_spPageContextInfo.serverRequestPath);
        //var page = spo.web.getFileByServerRelativeUrl("/sites/acctmgmt/DemroTest-DemroTestProjectII/SitePages/Home_copy(5).aspx");
        
        var dfd = new jQuery.Deferred();
        
        console.log("Loading new context");
        this.ctx = new SP.ClientContext(siteAbsoluteUrl);
        this.web = spo.ctx.get_web();
        spo.ctx.load(this.web);
        spo.ctx.executeQueryAsync(function() {
            
            var page = spo.web.getFileByServerRelativeUrl(pageRelativeUrl);
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
                                dfd.resolve();
                            });    
                        });
                        
                        
                    });
                });
            });
        });
        
        return dfd.promise();
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
    //             + '               network: \'company.com\'  }); '
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
 

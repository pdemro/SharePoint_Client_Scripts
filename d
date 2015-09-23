[1mdiff --git a/replace_webpart.js b/replace_webpart.js[m
[1mindex b85bcca..8b992c2 100644[m
[1m--- a/replace_webpart.js[m
[1m+++ b/replace_webpart.js[m
[36m@@ -105,7 +105,7 @@[m [mvar deleteKeyCommunciationsWebPart = function(ctx, limitedWebPartManager) {[m
 			[m
 			var index = -1;[m
 			jQuery.each(data.d.results, function(i,v) {[m
[31m-				if(v.WebPart.Title == "Key Communications"){[m
[32m+[m				[32mif(v.WebPart.Title == webPartTitle){[m
 					index = i;[m
 					return false;[m
 				}[m
[36m@@ -119,13 +119,13 @@[m [mvar deleteKeyCommunciationsWebPart = function(ctx, limitedWebPartManager) {[m
 					webPartsClient.get_item(index).deleteWebPart();[m
 					[m
 					ctx.executeQueryAsync(function() {[m
[31m-						console.log("deleted Key Communication web part");[m
[32m+[m						[32mconsole.log("deleted "+ webPartTitle + " web part");[m
 						deferred.resolve();[m
 					});[m
 				});[m
 			}[m
 			else {[m
[31m-				console.log("Did not find key communications web part");[m
[32m+[m				[32mconsole.log("Did not find " + webPartTitle + " web part");[m
 				deferred.resolve();[m
 			}[m
 		},[m
[36m@@ -183,17 +183,17 @@[m [mvar replaceKeyCommunicationsWebPart = function(clientContext){[m
 [m
 var onload = function() { [m
 	var competencyUrls = [[m
[31m-		//"https://chemours.sharepoint.com/teams/SHEEC_PORTAL1/auditing/",[m
[32m+[m		[32m"/teams/SHEEC_PORTAL1/auditing/",[m
 		"/teams/SHEEC_PORTAL1/data-systems-reporting",[m
[31m-		 //"/teams/SHEEC_PORTAL1/distribution",[m
[31m-		// "https://chemours.sharepoint.com/teams/SHEEC_PORTAL1/emergency-responses",[m
[31m-		// "https://chemours.sharepoint.com/teams/SHEEC_PORTAL1/environmental-stewardship",[m
[31m-		// "https://chemours.sharepoint.com/teams/SHEEC_PORTAL1/ergonomics",[m
[31m-		// "https://chemours.sharepoint.com/teams/SHEEC_PORTAL1/fire-safety",[m
[31m-		// "https://chemours.sharepoint.com/teams/SHEEC_PORTAL1/occupational-health",[m
[31m-		// "https://chemours.sharepoint.com/teams/SHEEC_PORTAL1/occupational-medicine",[m
[31m-		// "https://chemours.sharepoint.com/teams/SHEEC_PORTAL1/psm",[m
[31m-		// "https://chemours.sharepoint.com/teams/SHEEC_PORTAL1/workplace-safety"[m
[32m+[m		[32m"/teams/SHEEC_PORTAL1/distribution",[m
[32m+[m		[32m"/teams/SHEEC_PORTAL1/emergency-responses",[m
[32m+[m		[32m"/teams/SHEEC_PORTAL1/environmental-stewardship",[m
[32m+[m		[32m"/teams/SHEEC_PORTAL1/ergonomics",[m
[32m+[m		[32m"/teams/SHEEC_PORTAL1/fire-safety",[m
[32m+[m		[32m"/teams/SHEEC_PORTAL1/occupational-health",[m
[32m+[m		[32m"/teams/SHEEC_PORTAL1/occupational-medicine",[m
[32m+[m		[32m"/teams/SHEEC_PORTAL1/psm",[m
[32m+[m		[32m"/teams/SHEEC_PORTAL1/workplace-safety"[m
 		][m
 		[m
 	for(var i = 0; i < competencyUrls.length; i++) {[m
[36m@@ -207,7 +207,7 @@[m [mvar onload = function() {[m
 };[m
 [m
 [m
[31m-[m
[32m+[m[32mvar webPartTitle = "Key Communications";[m
 [m
 // [m
 // SP.SOD.registerSod("jquery", "https://ajax.googleapis.com/ajax/libs/jquery/2.1.3/jquery.min.js");[m

// Batch Check-in scripts for SPTools
// Copyright (c)2012 Josh McCarty

// Prevent console.log errors in IE
window.log = function() {
	log.history = log.history || [];
	log.history.push( arguments );
	if ( this.console ) {
		console.log( Array.prototype.slice.call( arguments ) );
	}
};

function checkInDocuments( performCheckIn ) {
	
	// Get form field values
	var sourceWebUrl = $( '#source-web-url' ).val();
	var sourceListName = $( '#source-list-name' ).val();
	var camlQuery = $( '#caml-query' ).val();
	var serverUrl = $( '#server-url' ).val();
	var checkInType = $( 'input[name="check-in-type"]' ).val();
	
	// Check that required fields are filled out
	if ( sourceWebUrl === '' || sourceListName === '' || camlQuery === '' || serverUrl === '' ) {
		
		// Todo: more robust validation and messages
		alert( 'Please enter all required information for the source library.' );
		return false;
	}
	
	// Indicate that something is happening
	$( '#processing' ).remove();
	$( '#update-button' ).attr( 'disabled', 'disabled' ).parent().after( '<div id="processing" class="alert">Processing&hellip;</div>' );
	$( '#results-table' ).find( 'tbody' ).html( '' );	
	
	// Cache the processing message
	var processing = $( '#processing' );
	
	// Cache variable to store count of items processed
	var currentCount = 0;
	
	$().SPServices({
		operation: "GetListItems",
		async: true,
		webURL: sourceWebUrl,
		listName: sourceListName,
		CAMLQuery: camlQuery,
		CAMLRowLimit: 0,
		CAMLViewFields: '<ViewFields Properties="True" />',
		CAMLQueryOptions: '<QueryOptions><ViewAttributes Scope="RecursiveAll" IncludeRootFolder="True" /></QueryOptions>',
		completefunc: function( xData, Status ) {
		
			// Check to see if any items match the CAML query in the specified Web URL and List Name
			if ( $( xData.responseXML ).SPFilterNode( 'z:row' ).length !== 0 ) {
			
				var recordTotal = $( xData.responseXML ).SPFilterNode( 'z:row' ).length;
				
				// We have items to process
				$( xData.responseXML ).SPFilterNode( 'z:row' ).each( function() {
				
					var id = $( this ).attr( 'ows_ID' );
					var itemURL = serverUrl + "/" + $( this ).attr( 'ows_FileRef' ).split( ';#' )[1];
					var sourceUrl = "/" + $( this ).attr( 'ows_FileDirRef' ).split( ';#' )[1] + "/Forms/DispForm.aspx?ID=" + id;
				
					// Is this just a preview request?
					if ( performCheckIn !== true ) {
						
						// Indicate progress
						currentCount++;
						if ( recordTotal !== currentCount ) {
							$( processing ).html( 'Processing&hellip;' + currentCount + ' of ' + recordTotal );
						}
						else {
							$( processing ).addClass( 'alert-success' ).html( 'Preview Complete! ' + recordTotal + ' items processed.' );
							$( '#update-button' ).removeAttr( 'disabled' );
						}
						
						$( '#results-table' ).find( 'tbody' ).append( '<tr><td><a href="' + sourceUrl + '">' + sourceUrl + '</a></td><td>This document will be checked in.</td></tr>' );
					}
					else {
						
						var today = new Date();
						
						// Check in the document using the itemURL, pass along the id to update the appropriate table row
						$().SPServices({
							operation: "CheckInFile",
							async: true,
							pageUrl: itemURL,
							comment: "This file was batch checked in on " + today.getFullYear() + "-" + ( today.getMonth() + 1 ) + "-" + today.getDate(),
							CheckinType: "1", // Publish a major version
							completefunc: function ( xData, Status ) {	
								
								var resultClass = "";
								var resultText = "";
								
								if ( Status !== "success" ) {
								
									// We'll log the responseXML for debugging if there is an error
									console.log( "Error for " + itemURL );
									resultClass = "alert-error";
									resultText = "Error";
								}
								else {
								
									resultText = "Success"
								}
								
								// Indicate progress
								currentCount++;
								if ( recordTotal !== currentCount ) {
									$( processing ).html( 'Processing&hellip;' + currentCount + ' of ' + recordTotal );
								}
								else {
									$( processing ).addClass( 'alert-success' ).html( 'Updates Complete! ' + recordTotal + ' items processed.' );
									$( '#update-button' ).removeAttr( 'disabled' );
								}
								
								// Add to results table
								$( '#results-table' ).find( 'tbody' ).append( '<tr id="item-' + id + '"><td><a href="' + sourceUrl + '">' + sourceUrl + '</a></td><td><span class="nowrap">Check In: </span><span class="' + resultClass + '">' + resultText + '</span><br />See console log for details.</td></tr>' );
								
								console.log( $( xData.responseXML ) );
							}
						});
					}					
				});
			}
			else {
				// We don't have any items to process
				var message = "There are no documents that match the specified CAML Query, Web URL, and Library Name.";
				console.log( message );
				$( processing ).addClass( 'alert-error' ).html( message );
				$( '#update-button' ).removeAttr( 'disabled' );
				$( '#results-table' ).find( 'tbody' ).append( '<tr><td colspan="2">' + message + '</td></tr>' );
			}
		}
	});	
}

$( document ).ready( function() {
	$( '#batch-form' ).submit( function() {
		checkInDocuments( true );
		return false;
	});
	$( '#preview-button' ).click( function() {
		checkInDocuments( false );
		return false;
	});
});
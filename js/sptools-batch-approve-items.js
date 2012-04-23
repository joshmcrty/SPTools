// Batch Approve Items scripts for SPTools
// Copyright (c)2012 Josh McCarty

// Prevent console.log errors in IE
window.log = function() {
	log.history = log.history || [];
	log.history.push( arguments );
	if ( this.console ) {
		console.log( Array.prototype.slice.call( arguments ) );
	}
};

function approveListItems( performApproval ) {

	// Get form field values
	var sourceWebUrl = $( '#source-web-url' ).val();
	var sourceListName = $( '#source-list-name' ).val();
	var camlQuery = $( '#caml-query' ).val();
	
	// Check that required fields are filled out
	if ( sourceWebUrl === '' || sourceListName === '' || camlQuery === '' ) {
		
		// Todo: more robust validation and messages
		alert( 'Please enter all required information for the source list.' );
		return false;
	}
	
	// Todo: check that list actually requires content approval.
	
	// Indicate that something is happening
	$( '#processing' ).remove();
	$( '#update-button' ).attr( 'disabled', 'disabled' ).parent().after( '<div id="processing" class="alert">Processing&hellip;</div>' );
	$( '#results-table' ).find( 'tbody' ).html( '' );
	
	// Cache the processing message
	var processing = $( '#processing' );
	
	// Cache variable to store count of items processed
	var currentCount = 0;
	
	// Get each source list item
	$().SPServices({
		operation: 'GetListItems',
		async: true,
		listName: sourceListName,
		webURL: sourceWebUrl,
		CAMLRowLimit: 0,
		CAMLQuery: camlQuery,
		CAMLViewFields: '<ViewFields Properties="True" />',
		completefunc: function( xData, Status ) {
			
			// Check to see if any items match the CAML query in the specified Web URL and List Name
			if ( $( xData.responseXML ).SPFilterNode( 'z:row' ).length !== 0 ) {
			
				var recordTotal = $( xData.responseXML ).SPFilterNode( 'z:row' ).length;
			
				$( xData.responseXML ).SPFilterNode( 'z:row' ).each( function( index ) {
					
					var sourceID = $( this ).attr( 'ows_ID' );
					var sourceDir = $( this ).attr( 'ows_FileDirRef' ).split( ';#' )[1];
					var sourceUrl = '/' + sourceDir + '/DispForm.aspx?ID=' + sourceID;
					
					// Is this just a preview request?
					if ( performApproval !== true ) {
						
						// Indicate progress
						currentCount++;
						if ( recordTotal !== currentCount ) {
							$( processing ).html( 'Processing&hellip;' + currentCount + ' of ' + recordTotal );
						}
						else {
							$( processing ).addClass( 'alert-success' ).html( 'Preview Complete! ' + recordTotal + ' items processed.' );
							$( '#update-button' ).removeAttr( 'disabled' );
						}
						
						$( '#results-table' ).find( 'tbody' ).append( '<tr><td><a href="' + sourceUrl + '">' + sourceUrl + '</a></td><td>This item will be approved.</td></tr>' );
					}
					else {
					
						var sourceID = $( this ).attr( 'ows_ID' );
						var sourceDir = $( this ).attr( 'ows_FileDirRef' ).split( ';#' )[1];
						var sourceUrl = '/' + sourceDir + '/DispForm.aspx?ID=' + sourceID;
						
						
						// Cache this node so we can retrieve the attributes for it
						var $node = $( this );
						
						// Update the list item using the batchCmd variable
						$().SPServices({
							operation: 'UpdateListItems',
							async: true,
							listName: sourceListName,
							webURL: sourceWebUrl,
							ID: sourceID,
							batchCmd: "Moderate",
							valuepairs: [["_ModerationStatus", "0"]],
							completefunc: function( xData, Status ) {
							
								var resultClass = "";
								var resultText = "";
								
								if ( Status !== "success" ) {
								
									// We'll log the responseXML for debugging if there is an error
									console.log( "Error for " + sourceUrl );
									console.log( $( xData.responseXML ) );				
									resultClass = "error";
									resultText = "Error";
								}
								else {
									var errorCode = $( xData.responseXML ).find( 'ErrorCode' ).text();
									if ( errorCode !== "0x00000000" ) {
										resultText = "Error";
									}
									else {
										resultText = "Success";
									}
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
								$( '#results-table' ).find( 'tbody' ).append( '<tr class="' + resultClass + '" id="item-' + sourceID + '"><td><a href="' + sourceUrl + '">View Approved Item</a></td><td><span class="nowrap">Approval: </span> ' + resultText + '<br />See console log for details.</td></tr>' );
								
								console.log( $( xData.responseXML ) );
							}
						});
					}
				});
			}
			else {
			
				// We don't have any items to process
				var message = "There are no source list items that match the specified CAML Query, Web URL, and List Name.";
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
		approveListItems( true );
		return false;
	});
	$( '#preview-button' ).click( function() {
		approveListItems( false );
		return false;
	});
});
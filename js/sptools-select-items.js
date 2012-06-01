// Select Items scripts for SPTools
// Copyright (c)2012 Josh McCarty

// Prevent console.log errors in IE
window.log = function() {
	log.history = log.history || [];
	log.history.push( arguments );
	if ( this.console ) {
		console.log( Array.prototype.slice.call( arguments ) );
	}
};

function selectListItems( processItems ) {

	// Get form field values
	var sourceWebUrl = $( '#source-web-url' ).val();
	var sourceListName = $( '#source-list-name' ).val();
	var camlQuery = $( '#caml-query' ).val();
	var additionalFilter = $( 'input[name="additional-filter"]:checked' ).val();
	var camlOptions = "";
	var camlRowLimit = 0;
	
	if ( additionalFilter === "limit-number" ) {
		camlRowLimit = Number( $( '#additional-filter-limit-number' ).val() );
	}
	
	// Check that required fields are filled out
	if ( processItems === false ) {
		if ( sourceWebUrl === '' || sourceListName === '' || camlQuery === '' ) {
			
			// Todo: more robust validation and messages
			alert( 'Please enter all required information for the source list.' );
			return false;
		}
	}
	else {
		
		// Check that required fields are filled out
		if ( sourceWebUrl === '' || sourceListName === '' || camlQuery === '' ) {
			
			// Todo: more robust validation and messages
			alert( 'Please enter all required information for the source list.' );
			return false;
		}
	}
	
	// Indicate that something is happening
	$( '#processing' ).remove();
	$( '#update-button' ).attr( 'disabled', 'disabled' ).parent().after( '<div id="processing" class="alert">Processing&hellip;Please be patient as this may take several minutes and your browser may become unresponsive.</div>' );
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
		CAMLRowLimit: camlRowLimit,
		CAMLQuery: camlQuery,
		CAMLViewFields: '<ViewFields Properties="True" />',
		completefunc: function( xData, Status ) {
			
			// Check to see if any items match the CAML query in the specified Web URL and List Name
			if ( $( xData.responseXML ).SPFilterNode( 'z:row' ).length !== 0 ) {
			
				var recordTotal = $( xData.responseXML ).SPFilterNode( 'z:row' ).length;
			
				// Is this just a preview request?
				if ( processItems !== true ) {
				
					$( xData.responseXML ).SPFilterNode( 'z:row' ).each( function( index ) {
						
						var sourceID = $( this ).attr( 'ows_ID' );
						var sourceDir = $( this ).attr( 'ows_FileDirRef' ).split( ';#' )[1];
						var sourceUrl = '/' + sourceDir + '/DispForm.aspx?ID=' + sourceID;
							
						// Indicate progress
						currentCount++;
						if ( recordTotal !== currentCount ) {
							$( processing ).html( 'Processing&hellip;' + currentCount + ' of ' + recordTotal );
						}
						else {
							$( processing ).addClass( 'alert-success' ).html( 'Preview Complete! ' + recordTotal + ' items processed.' );
							$( '#update-button' ).removeAttr( 'disabled' );
						}
						
						$( '#results-table' ).find( 'tbody' ).append( '<tr><td><a href="' + sourceUrl + '">' + sourceUrl + '</a></td><td>This item will be additionally filtered.</td></tr>' );
					});
				}
				else {
				
					if ( additionalFilter === "random-select" ) {
						
						var randomNumber = Math.floor( Math.random() * recordTotal );
						var listItems = $( xData.responseXML ).SPFilterNode( 'z:row' );
						var randomItem = listItems[randomNumber];
						
						$( randomItem ).each( function( index ) {
							
							var sourceID = $( this ).attr( 'ows_ID' );
							var sourceDir = $( this ).attr( 'ows_FileDirRef' ).split( ';#' )[1];
							var sourceUrl = '/' + sourceDir + '/DispForm.aspx?ID=' + sourceID;
								
							// Indicate progress
							$( processing ).addClass( 'alert-success' ).html( 'Selection Complete! ' + recordTotal + ' items processed and filtered (not all items will be displayed below depending on the filters).' );
							$( '#update-button' ).removeAttr( 'disabled' );
							
							// Add to results table
							$( '#results-table' ).find( 'tbody' ).append( '<tr class="" id="item-' + sourceID + '"><td><a href="' + sourceUrl + '">View Selected Item</a></td><td>See console log for details.</td></tr>' );
							
							console.log( $( this ) );
							
						});
					}
					else {
						
						$( xData.responseXML ).SPFilterNode( 'z:row' ).each( function() {
							
							var sourceID = $( this ).attr( 'ows_ID' );
							var sourceDir = $( this ).attr( 'ows_FileDirRef' ).split( ';#' )[1];
							var sourceUrl = '/' + sourceDir + '/DispForm.aspx?ID=' + sourceID;
								
							// Indicate progress
							currentCount++;
							if ( recordTotal !== currentCount ) {
								$( processing ).html( 'Processing&hellip;' + currentCount + ' of ' + recordTotal );
							}
							else {
								$( processing ).addClass( 'alert-success' ).html( 'Preview Complete! ' + recordTotal + ' items processed.' );
								$( '#update-button' ).removeAttr( 'disabled' );
							}
							
							$( processing ).addClass( 'alert-success' ).html( 'Selection Complete! ' + recordTotal + ' items processed and filtered (not all items will be displayed below depending on the filters).' );
							$( '#update-button' ).removeAttr( 'disabled' );
							
							// Add to results table
							$( '#results-table' ).find( 'tbody' ).append( '<tr class="" id="item-' + sourceID + '"><td><a href="' + sourceUrl + '">View Selected Item</a></td><td>See console log for details.</td></tr>' );
							
							console.log( $( this ) );
						});
					}
				}
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
		selectListItems( true );
		return false;
	});
	$( '#preview-button' ).click( function() {
		selectListItems( false );
		return false;
	});
});
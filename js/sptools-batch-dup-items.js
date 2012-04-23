// Batch Duplicate Items scripts for SPTools
// Copyright (c)2012 Josh McCarty

// Prevent console.log errors in IE
window.log = function() {
	log.history = log.history || [];
	log.history.push( arguments );
	if ( this.console ) {
		console.log( Array.prototype.slice.call( arguments ) );
	}
};

function getListFields( listName, webURL, defaultValue ) {

	// Cache the array to store the fields names in
	var html = '<select><option value="">Select&hellip;</option>';
	
	$().SPServices({
		operation: 'GetList',
		async: false,
		listName: listName,
		webURL: webURL,
		completefunc: function( xData, Status ) {
			
			// Add each field to the array of fields
			$( xData.responseXML ).SPFilterNode( 'Field' ).each( function() {				
				var staticName = $( this ).attr( 'StaticName' );
				var displayName = $( this ).attr( 'DisplayName' );
				var baseColumn = $( this ).attr( 'FromBaseType' );
				var requiredColumn = $( this ).attr( 'Required' );
				
				// Indicate if a column is required with an asterisk
				var requiredColumnIndicator = "";
				if ( requiredColumn === "TRUE" ) {
					requiredColumnIndicator = "*";
				}
				
				if ( typeof staticName !== "undefined" && typeof displayName !== "undefined" && ( baseColumn !== "TRUE" || requiredColumn === "TRUE" || staticName === "Title" ) ) {
					if ( staticName === defaultValue ) {
						html += '<option value="' + staticName + '" selected="selected">' + displayName + requiredColumnIndicator + '</option>';
					}
					else {
						html += '<option value="' + staticName + '">' + displayName + '</option>';
					}
				}
			});
			html += '</select>';
		}
	});
	return html;
}

function addField( defaultValue ) {
	
	// Get form field values
	var sourceWebUrl = $( '#source-web-url' ).val();
	var sourceListName = $( '#source-list-name' ).val();
	var camlQuery = $( '#caml-query' ).val();
	var destinationWebUrl = $( '#destination-web-url' ).val();
	var destinationListName = $( '#destination-list-name' ).val();
	
	if ( sourceWebUrl === '' || sourceListName === '' || camlQuery === '' || destinationWebUrl === '' || destinationListName === '' ) {
		
		// Todo: more robust validation and messages
		alert( 'Please enter all required information for the source and destination lists.' );
		return false;
	}
	
	// Get the list of fields for the source and destination lists
	var sourceFieldList = getListFields( sourceListName, sourceWebUrl, defaultValue );
	var destinationFieldList = getListFields( destinationListName, destinationWebUrl, defaultValue );
	
	// Cache the table body element
	var tableBody = $( '#field-list-table' ).find( 'tbody' );
	
	// Add a table row with the source and destination list fields and dropdown menus
	$( tableBody ).append( '<tr><td class="source-field-select">' + sourceFieldList + '<br /><a href="#remove-field" class="remove-field" title="Remove this field">Remove</a></td><td class="destination-field-select">' + destinationFieldList + '</td></tr>' );
		
	// Automatically switch focus to the next field; automatically add another row when the destination field is chosen
	$( tableBody ).find( 'tr:last-child' ).find( '.source-field-select' ).find( 'select' ).focus().change( function() {
		var sourceText = $( this ).find( 'option:selected' ).text();
		console.log( sourceText );
		var destinationSelect = $( this ).closest( 'tr' ).find( '.destination-field-select' ).find( 'select' );
		$( destinationSelect ).focus().find( 'option' ).each( function() {
			
			// If there is a field with an identical display name, pre-select it
			if ( $( this ).text() === sourceText ) {
				var destinationValue = $( this ).attr( 'value' );
				$( destinationSelect ).val( destinationValue );
			}
		});
		
	});
	$( '.remove-field' ).click( function() {
		$( this ).closest( 'tr' ).remove();
		return false;
	});
}

function duplicateListItems( performCopy ) {

	// Get form field values
	var sourceWebUrl = $( '#source-web-url' ).val();
	var sourceListName = $( '#source-list-name' ).val();
	var destinationWebUrl = $( '#destination-web-url' ).val();
	var destinationListName = $( '#destination-list-name' ).val();
	var camlQuery = $( '#caml-query' ).val();
	
	// Check that required fields are filled out
	if ( performCopy === false ) {
		if ( sourceWebUrl === '' || sourceListName === '' || camlQuery === '' ) {
			
			// Todo: more robust validation and messages
			alert( 'Please enter all required information for the source list.' );
			return false;
		}		
	}
	else {
		if ( sourceWebUrl === '' || sourceListName === '' || camlQuery === '' || destinationWebUrl === '' || destinationListName === '' ) {
			
			// Todo: more robust validation and messages
			alert( 'Please enter all required information for the source and destination lists.' );
			return false;
		}
		
		// Check for source field mapping
		else if ( $( '#field-list-table' ).find( 'tbody' ).find( 'tr' ).find( '.source-field-select' ).length === 0 ) {
			
			// Todo: more robust validation and messaging.
			alert( 'You have not chosen any fields to map. Please map all fields that you want to copy.' );
			return false;
		}
	}
	
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
					var destinationUrl = "";
					
					// Is this just a preview request?
					if ( performCopy !== true ) {
						
						// Indicate progress
						currentCount++;
						if ( recordTotal !== currentCount ) {
							$( processing ).html( 'Processing&hellip;' + currentCount + ' of ' + recordTotal );
						}
						else {
							$( processing ).addClass( 'alert-success' ).html( 'Complete! ' + recordTotal + ' items processed.' );
							$( '#update-button' ).removeAttr( 'disabled' );
						}
						
						$( '#results-table' ).find( 'tbody' ).append( '<tr><td><a href="' + sourceUrl + '">' + sourceUrl + '</a></td><td>This item will be copied to ' + destinationListName + '</td></tr>' );
					}
					else {
					
						var sourceID = $( this ).attr( 'ows_ID' );
						var sourceDir = $( this ).attr( 'ows_FileDirRef' ).split( ';#' )[1];
						var sourceUrl = '/' + sourceDir + '/DispForm.aspx?ID=' + sourceID;
						var destinationUrl = "";
						
						// Create the batchCmd variable that will be used to create the new list item
						var batchCmd = '<Batch OnError="Continue"><Method ID="1" Cmd="New">';
						
						// Cache this node so we can retrieve the attributes for it
						var $node = $( this );
						
						// Get the value of the selected source fields and add them to the batchCmd variable for the destination fields of the new list item
						$( '#field-list-table' ).find( 'tbody' ).find( 'tr' ).each( function() {
							var sourceField = $( this ).find( '.source-field-select' ).find( 'select' ).val();
							var destinationField = $( this ).find( '.destination-field-select' ).find( 'select' ).val();
							
							// If the source field is not empty, add it to the batchCmd
							if ( typeof $node.attr( 'ows_' + sourceField ) !== "undefined" ) {
								batchCmd += '<Field Name="' + destinationField + '"><![CDATA[' + $node.attr( 'ows_' + sourceField ) + ']]></Field>';
							}
						});
						
						// Close the batchCmd variable
						batchCmd += '</Method></Batch>';
						
						// Create a new list item on the destination list using the batchCmd variable
						$().SPServices({
							operation: 'UpdateListItems',
							async: true,
							listName: destinationListName,
							webURL: destinationWebUrl,
							updates: batchCmd,
							completefunc: function( xData, Status ) {
								
								var resultClass = "";
								var resultText = "";
								
								if ( Status !== "success" ) {
								
									// We'll log the responseXML for debugging if there is an error
									console.log( "Error for " + itemURL );
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
									$( xData.responseXML ).SPFilterNode( 'z:row' ).each( function() {
										var destinationID = $( this ).attr( 'ows_ID' );
										var destinationDir = $( this ).attr( 'ows_FileDirRef' ).split( ';#' )[1];
										destinationUrl = '/' + destinationDir + '/DispForm.aspx?ID=' + destinationID;
									});
								}
								
								// Indicate progress
								currentCount++;
								if ( recordTotal !== currentCount ) {
									$( processing ).html( 'Processing&hellip;' + currentCount + ' of ' + recordTotal );
								}
								else {
									$( processing ).addClass( 'alert-success' ).html( 'Complete! ' + recordTotal + ' items processed.' );
									$( '#update-button' ).removeAttr( 'disabled' );
								}
								
								// Add to results table
								$( '#results-table' ).find( 'tbody' ).append( '<tr class="' + resultClass + '" id="item-' + sourceID + '"><td><a href="' + sourceUrl + '">Source Item</a> | <a href="' + destinationUrl + '">Destination Item</a></td><td><span class="nowrap">Duplication: </span> ' + resultText + '<br />See console log for details.</td></tr>' );
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
		duplicateListItems( true );
		return false;
	});
	$( '#reset-button' ).click( function() {
		$( '#field-list-table' ).find( 'tbody' ).html( '' );
		return false;
	});
	$( '#preview-button' ).click( function() {
		duplicateListItems( false );
		return false;
	});
	$( '#add-field-button' ).click( function() {
		addField();
		return false;
	});
});
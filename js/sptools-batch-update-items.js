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

	// Cache the variable to store the fields names in
	var html = '<select><option value="" selected="selected">Select&hellip;</option>';
	
	$().SPServices({
		operation: 'GetList',
		async: false,
		listName: listName,
		webURL: webURL,
		completefunc: function( xData, Status ) {
		
			console.log( $( xData.responseXML ) );
			
			// Add each field to the array of fields
			$( xData.responseXML ).SPFilterNode( 'Field' ).each( function() {				
				var staticName = $( this ).attr( 'StaticName' );
				var displayName = $( this ).attr( 'DisplayName' );
				var baseColumn = $( this ).attr( 'FromBaseType' );
				var requiredColumn = $( this ).attr( 'Required' );
				var hiddenColumn = $( this ).attr( 'Hidden' );
				var typeColumn = $( this ).attr( 'Type' );
				var readOnlyColumn = $( this ).attr( 'ReadOnly' );
				
				// Indicate if a column is required with an asterisk
				var requiredColumnIndicator = "";
				if ( requiredColumn === "TRUE" ) {
					requiredColumnIndicator = "*";
				}

				
				if ( typeof staticName !== "undefined" && typeof displayName !== "undefined" && hiddenColumn !== "TRUE" && typeColumn !== "Computed" && readOnlyColumn !== "TRUE" ) {

					if ( staticName === defaultValue ) {
						html += '<option value="' + staticName + '" selected="selected" data-hidden="' + hiddenColumn + '">' + displayName + requiredColumnIndicator + '</option>';
					}
					else {
						html += '<option value="' + staticName + '" data-hidden="' + hiddenColumn + '">' + displayName + requiredColumnIndicator + '</option>';
					}
				}
			});
			html += '</select>';
		}
	});
	return html;
}

function getListUpdates( listName, webURL, fieldName ) {
	
	var now = new Date();
	var uniqueID = now.valueOf();
	
	// Create the dropdown for the field.
	var html = '<label for="' + uniqueID + '">Select the type of update to perform:</label> <select id="' + uniqueID + '" class="source-field-update-type"><option value="">Select&hellip;</option><option value="remove-lookup">Remove Lookup Item</option><option value="set-value">Set Value</option><option value="append-value">Append Value</option><option value="prepend-value">Prepend Value</option><option value="clear-value">Clear Value</option>';
	
	html += '</select>';
	
	return html;
}

function getCurrentValue( sourceListName, sourceWebUrl, sourceField, sourceID ) {
	
	var fieldValue = "";
	
	// Get each source list item
	$().SPServices({
		operation: 'GetListItems',
		async: false,
		listName: sourceListName,
		webURL: sourceWebUrl,
		CAMLRowLimit: 0,
		CAMLQuery: '<Query><Where><Eq><FieldRef Name="ID" /><Value Type="Integer">' + sourceID + '</Value></Eq></Where></Query>',
		CAMLViewFields: '<ViewFields Properties="True" />',
		completefunc: function( xData, Status ) {
			
			$( xData.responseXML ).SPFilterNode( 'z:row' ).each( function() {
				fieldValue = $( this ).attr( 'ows_' + sourceField );
			});
		}
	});
	return fieldValue;
}

function addField( defaultValue ) {
	
	// Get form field values
	var sourceWebUrl = $( '#source-web-url' ).val();
	var sourceListName = $( '#source-list-name' ).val();
	var camlQuery = $( '#caml-query' ).val();
	
	if ( sourceWebUrl === '' || sourceListName === '' || camlQuery === '' ) {
		
		// Todo: more robust validation and messages
		alert( 'Please enter all required information for the source list.' );
		return false;
	}
	
	// Get the list of fields for the source list
	var sourceFieldList = getListFields( sourceListName, sourceWebUrl, defaultValue );
	
	// Cache the table body element
	var tableBody = $( '#field-list-table' ).find( 'tbody' );
	
	// Add a table row with the source list field and dropdown menu
	$( tableBody ).append( '<tr><td class="source-field-select">' + sourceFieldList + '<br /><a href="#remove-field" class="remove-field" title="Remove this field">Remove</a></td><td class="source-field-update">&larr; Select a field to apply updates to.</td></tr>' );
		
	// Automatically switch focus to the next field; automatically add another row when the destination field is chosen
	$( tableBody ).find( 'tr:last-child' ).find( '.source-field-select' ).find( 'select' ).focus().change( function() {
		var fieldName = $( this ).val();
		var sourceFieldUpdates = getListUpdates( sourceListName, sourceWebUrl, fieldName );
		$( tableBody ).find( 'tr:last-child' ).find( '.source-field-update' ).html( sourceFieldUpdates );
		$( tableBody ).find( 'tr:last-child' ).find( '.source-field-update' ).find( 'select' ).change( function() {
			
			$( this ).parent().find( '.update-wrapper' ).remove();
			
			var updateType = $( this ).val();
			switch( updateType ) {
			
				case "remove-lookup":
					
					var now = new Date();
					var uniqueID = now.valueOf();
					$( this ).parent().append( '<div class="update-wrapper"><label for="' + uniqueID + '">Enter the lookup item to remove:</label> <input id="' + uniqueID + '" class="source-field-update-value" type="text" placeholder="1;#Item 1" /></div>' );
					
					break;
					
				case "set-value":
				
					var now = new Date();
					var uniqueID = now.valueOf();
					$( this ).parent().append( '<div class="update-wrapper"><label for="' + uniqueID + '">Enter the new value for this field in the proper format:</label> <input id="' + uniqueID + '" class="source-field-update-value" type="text" placeholder="Hello World" /></div>' );
					
					break;
					
				case "append-value":
				
					var now = new Date();
					var uniqueID = now.valueOf();
					$( this ).parent().append( '<div class="update-wrapper"><label for="' + uniqueID + '">Enter the value to append to the current value:</label> <input id="' + uniqueID + '" class="source-field-update-value" type="text" placeholder="&hellip;and this too" /></div>' );
					
					break;
					
				case "prepend-value":
				
					var now = new Date();
					var uniqueID = now.valueOf();
					$( this ).parent().append( '<div class="update-wrapper"><label for="' + uniqueID + '">Enter the value to prepend to the current value:</label> <input id="' + uniqueID + '" class="source-field-update-value" type="text" placeholder="This first&hellip;" /></div>' );
					
					break;
				
				case "clear-value":
				
					var now = new Date();
					var uniqueID = now.valueOf();
					$( this ).parent().append( '<div class="update-wrapper">This field&rsquo;s value will be cleared (it will be empty).</div>' );
					
					break;
					
				default:
					
					// Do nothing
			}
		});
	});
	$( '.remove-field' ).click( function() {
		$( this ).closest( 'tr' ).remove();
		return false;
	});
}

function updateListItems( performUpdate ) {

	// Get form field values
	var sourceWebUrl = $( '#source-web-url' ).val();
	var sourceListName = $( '#source-list-name' ).val();
	var camlQuery = $( '#caml-query' ).val();
	
	// Check that required fields are filled out
	if ( performUpdate === false ) {
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
		
		// Check for source field mapping
		else if ( $( '#field-list-table' ).find( 'tbody' ).find( 'tr' ).find( '.source-field-select' ).length === 0 ) {
			
			// Todo: more robust validation and messaging.
			alert( 'You have not chosen any fields to update. Please select all fields that you want to update.' );
			return false;
		}
		else if ( $( '#field-list-table' ).find( 'tbody' ).find( 'tr' ).find( '.source-field-update-type' ).filter( function() { return $( this ).val() === ''; }).length !== 0 || $( '#field-list-table' ).find( 'tbody' ).find( 'tr' ).find( '.source-field-update-value' ).filter( function() { return $( this ).val() === ''; }).length !== 0 ) {
			alert( 'You have not selected an update type or new update value for all fields.' );
			return false;
		}
		
		// Todo: check for content approval and/or minor versions and warn user that items may need to be approved/published after the update
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
					if ( performUpdate !== true ) {
						
						// Indicate progress
						currentCount++;
						if ( recordTotal !== currentCount ) {
							$( processing ).html( 'Processing&hellip;' + currentCount + ' of ' + recordTotal );
						}
						else {
							$( processing ).addClass( 'alert-success' ).html( 'Preview Complete! ' + recordTotal + ' items processed.' );
							$( '#update-button' ).removeAttr( 'disabled' );
						}
						
						$( '#results-table' ).find( 'tbody' ).append( '<tr><td><a href="' + sourceUrl + '">' + sourceUrl + '</a></td><td>This item will be updated.</td></tr>' );
					}
					else {
					
						var sourceID = $( this ).attr( 'ows_ID' );
						var sourceDir = $( this ).attr( 'ows_FileDirRef' ).split( ';#' )[1];
						var sourceUrl = '/' + sourceDir + '/DispForm.aspx?ID=' + sourceID;
						
						// Create the batchCmd variable that will be used to update the list item
						var batchCmd = '<Batch OnError="Continue"><Method ID="1" Cmd="Update"><Field Name="ID">' + sourceID + '</Field>';
						
						// Cache this node so we can retrieve the attributes for it
						var $node = $( this );
						
						// Get the fields to update and how to update them
						$( '#field-list-table' ).find( 'tbody' ).find( 'tr' ).each( function() {
							var sourceField = $( this ).find( '.source-field-select' ).find( 'select' ).val();
							var sourceUpdateType = $( this ).find( '.source-field-update-type' ).val();
							var sourceUpdateValue = $( this ).find( '.source-field-update-value' ).val();
							
							// Get the current value that will be updated
							var sourceValue = getCurrentValue( sourceListName, sourceWebUrl, sourceField, sourceID );
							var newValue = "";
							
							// Calculate the new value based on the update type
							switch( sourceUpdateType ) {
								
								case "remove-lookup":
									
									newValue = sourceValue.split( sourceUpdateValue ).join( '' ).split( ';#;#' ).join( ';#' );
									
									break;
								
								case "set-value":
									
									newValue = sourceUpdateValue;
									
									break;
								
								case "append-value":
									
									newValue = sourceValue + sourceUpdateValue;
									
									break;
								
								case "prepend-value":
									
									newValue = sourceUpdateValue + sourceValue;
									
									break;
									
								case "remove-value":
									
									newValue = "";
									
									break;
									
								default:
								
									// Do nothing
							}
							batchCmd += '<Field Name="' + sourceField + '"><![CDATA[' + newValue + ']]></Field>';
						});
						
						// Close the batchCmd variable
						batchCmd += '</Method></Batch>';
						
						console.log( 'Batch command: ' + batchCmd );
						
						// Update the list item using the batchCmd variable
						$().SPServices({
							operation: 'UpdateListItems',
							async: true,
							listName: sourceListName,
							webURL: sourceWebUrl,
							updates: batchCmd,
							completefunc: function( xData, Status ) {
								
								var resultClass = "";
								var resultText = "";
								
								if ( Status !== "success" ) {
								
									// We'll log the responseXML for debugging if there is an error
									console.log( "Error for " + itemURL );
									console.log( $( xData.responseXML ) );				
									resultClass = "alert-error";
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
										var updatedID = $( this ).attr( 'ows_ID' );
										var updatedDir = $( this ).attr( 'ows_FileDirRef' ).split( ';#' )[1];
										updatedURL = '/' + updatedDir + '/DispForm.aspx?ID=' + updatedID;
									});
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
								$( '#results-table' ).find( 'tbody' ).append( '<tr id="item-' + sourceID + '"><td><a href="' + sourceUrl + '">View Updated Item</a></td><td><span class="nowrap">Update: </span><span class="' + resultClass + '">' + resultText + '</span><br />See console log for details.</td></tr>' );
								
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


function deleteListItems() {
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
					
					// Create the batchCmd variable that will be used to update the list item
					var batchCmd = '<Batch OnError="Continue"><Method ID="1" Cmd="Delete"><Field Name="ID">' + sourceID + '</Field></Method></Batch>';
					
					console.log( 'Batch command: ' + batchCmd );
					
					// Update the list item using the batchCmd variable
					$().SPServices({
						operation: 'UpdateListItems',
						async: true,
						listName: sourceListName,
						webURL: sourceWebUrl,
						updates: batchCmd,
						completefunc: function( xData, Status ) {
							
							var resultClass = "";
							var resultText = "";
							
							if ( Status !== "success" ) {
							
								// We'll log the responseXML for debugging if there is an error
								console.log( "Error for " + itemURL );
								console.log( $( xData.responseXML ) );				
								resultClass = "alert-error";
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
									var updatedID = $( this ).attr( 'ows_ID' );
									var updatedDir = $( this ).attr( 'ows_FileDirRef' ).split( ';#' )[1];
									updatedURL = '/' + updatedDir + '/DispForm.aspx?ID=' + updatedID;
								});
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
							$( '#results-table' ).find( 'tbody' ).append( '<tr id="item-' + sourceID + '"><td><a href="' + sourceUrl + '">Item Deleted</a></td><td><span class="nowrap">Update: </span><span class="' + resultClass + '">' + resultText + '</span><br />See console log for details.</td></tr>' );
							
							console.log( $( xData.responseXML ) );
						}
					});
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
	$( '#batch-form' ).submit( function( e ) {
		e.preventDefault();
		if ( $( '#delete-items' ).is( ':checked' ) ) {
			deleteListItems();
		}
		else {
			updateListItems( true );
		}
	});
	$( '#reset-button' ).click( function( e ) {
		e.preventDefault();
		$( '#field-list-table' ).find( 'tbody' ).html( '' );
	});
	$( '#preview-button' ).click( function( e ) {
		e.preventDefault();
		updateListItems( false );
	});
	$( '#add-field-button' ).click( function( e ) {
		e.preventDefault();
		addField();
	});
	$( '#delete-items' ).change( function() {
		console.log( 'changed' );
		if ( $( this ).is( ':checked' ) ) {
			$( '.spt-update-fields' ).addClass( 'sp-disabled-fieldset' ).prepend( '<div class="sp-disabled-overlay" />' );
		}
		else {
			$( '.spt-update-fields' ).removeClass( 'sp-disabled-fieldset' ).find( '.sp-disabled-overlay' ).remove();
		}
	});
});
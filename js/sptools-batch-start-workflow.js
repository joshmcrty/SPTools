// Batch Start Workflow scripts for SPTools
// Copyright (c)2012 Josh McCarty

// Prevent console.log errors in IE
window.log = function() {
	log.history = log.history || [];
	log.history.push( arguments );
	if ( this.console ) {
		console.log( Array.prototype.slice.call( arguments ) );
	}
};

function startWorkflow( performWorkflow ) {
	
	// Get form field values
	var sourceWebUrl = $( '#source-web-url' ).val();
	var sourceListName = $( '#source-list-name' ).val();
	var camlQuery = $( '#caml-query' ).val();
	var serverUrl = $( '#source-web-url' ).val();
	var workflowName = $( '#workflow-name' ).val();
	
	// Check that required fields are filled out
	if ( sourceWebUrl === '' || sourceListName === '' || camlQuery === '' || serverUrl === '' || workflowName === '' ) {
		
		// Todo: more robust validation and messages
		alert( 'Please enter all required information for the source list/library.' );
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
	
	var listType = "";
	
	// Get information about the list/library
	$().SPServices({
		operation: "GetList",
		listName: sourceListName,
		webURL: sourceWebUrl,
		completefunc: function( xData, Status ) {
			
			listType = $( xData.responseXML ).SPFilterNode( 'List' ).attr( 'ServerTemplate' );
		}
	});
	
	$().SPServices({
		operation: "GetListItems",
		async: true,
		webURL: sourceWebUrl,
		listName: sourceListName,
		CAMLQuery: camlQuery,
		CAMLRowLimit: 0,
		CAMLViewFields: '<ViewFields Properties="True" />',
		completefunc: function( xData, Status ) {
		
			// Check to see if any items match the CAML query in the specified Web URL and List Name
			if ( $( xData.responseXML ).SPFilterNode( 'z:row' ).length !== 0 ) {
			
				var recordTotal = $( xData.responseXML ).SPFilterNode( 'z:row' ).length;
				
				// We have items to process
				$( xData.responseXML ).SPFilterNode( 'z:row' ).each( function() {
				
					// Is this just a preview request?
					if ( performWorkflow !== true ) {
						
						// Don't attempt to initiate a workflow on a folder
						if ( $( this ).attr( 'ows_ContentType' ) !== "Folder" ) {
							var id = $( this ).attr( 'ows_ID' );
							var itemURL = serverUrl + "/" + $( this ).attr( 'ows_FileRef' ).split( ';#' )[1];
							var sourceUrl = "/" + $( this ).attr( 'ows_FileDirRef' ).split( ';#' )[1] + "/Forms/DispForm.aspx?ID=" + id;
							
							// Indicate progress
							currentCount++;
							if ( recordTotal !== currentCount ) {
								$( processing ).html( 'Processing&hellip;' + currentCount + ' of ' + recordTotal );
							}
							else {
								$( processing ).addClass( 'alert-success' ).html( 'Preview Complete! ' + recordTotal + ' items processed.' );
								$( '#update-button' ).removeAttr( 'disabled' );
							}
							
							$( '#results-table' ).find( 'tbody' ).append( '<tr><td><a href="' + sourceUrl + '">' + sourceUrl + '</a></td><td>This item will have the selected workflow initiated.</td></tr>' );
						}
					}
					else {
					
						// Don't attempt to initiate a workflow on a folder
						if ( $( this ).attr( 'ows_ContentType' ) !== "Folder" ) {
						
							var id = $( this ).attr( 'ows_ID' );
							var itemUrl = $( this ).attr( 'ows_EncodedAbsUrl' );
							var sourceUrl = "";
							var workflowGUID = "";
							
							if ( listType === "101" || listType === "109" || listType === "115" || listType === "119" || listType === "850" || listType === "2002" || listType ==="2003" ) {
							
								// This is a library with files
								sourceUrl = serverUrl + "/" + $( this ).attr( 'ows_FileDirRef' ).split( ';#' )[1] + "/Forms/DispForm.aspx?ID=" + id;
							}
							else {
								sourceUrl = serverUrl + "/" + $( this ).attr( 'ows_FileDirRef' ).split( ';#' )[1] + "/DispForm.aspx?ID=" + id;
							}
							
							// Now we can kick off the workflow. We'll get the templateId of the workflow using the selected workflow name.
							var workflowGUID = "";
							$().SPServices({
								operation: "GetTemplatesForItem",
								item: itemUrl,
								async: true,
								completefunc: function ( xData, Status ) {
								
									console.log( "GetTemplatesForItem" );
									console.log( $( xData.responseXML ) );
									
									$( xData.responseXML ).SPFilterNode ( 'WorkflowTemplate' ).each( function() {
										
										// Workflow name
										if ( $( this ).attr( 'Name' ) === workflowName ) {              
											var guid = $( this ).find( 'WorkflowTemplateIdSet' ).attr( 'TemplateId' );
											workflowGUID = "{" + guid + "}";
										}
										
										// Now we'll use SPServices to initiate the appropriate workflow by passing the itemUrl and the workflowGUID for the workflow.
										$().SPServices({
											operation: "StartWorkflow",
											item: itemUrl,
											async: true,
											templateId: workflowGUID,
											workflowParameters: "<root />",
											completefunc: function( xData, Status) {
												
												// Indicate progress
												currentCount++;
												if ( recordTotal !== currentCount ) {
													$( processing ).html( 'Processing&hellip;' + currentCount + ' of ' + recordTotal );
												}
												else {
													$( processing ).addClass( 'alert-success' ).html( 'Preview Complete! ' + recordTotal + ' items processed.' );
													$( '#update-button' ).removeAttr( 'disabled' );
												}
												
												var resultClass = "";
												var resultText = "The workflow has been started for this item.";
												
												if ( Status !== "success" ) {
													resultClass = "alert-error";
													resultText = "There was an error initiating the workflow.";
												}
												
												// Now that the workflow has been started, we'll add this item to the results table.
												$( '#results-table' ).find( 'tbody' ).append( '<tr><td><a href="' + sourceUrl + '">' + sourceUrl + '</a></td><td><span class="' + resultClass + '">' + resultText + '</span><br />See console log for details.</td></tr>' );
												
											}
										});
									});
								}
							});
						}
					}					
				});
			}
			else {
				// We don't have any items to process
				var message = "There are no documents that match the specified CAML Query, Web URL, and List/Library Name.";
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
		startWorkflow( true );
		return false;
	});
	$( '#preview-button' ).click( function() {
		startWorkflow( false );
		return false;
	});
});
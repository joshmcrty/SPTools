SPTools
=======

A collection of GUI tools for interacting with SharePoint&#39;s web services to perform batch operations.

These are still in development. I created them to serve my own needs, so their functionality is limited to what I've needed in my day job. I will be adding more functionality and cleaning up the UI as time permits.

General Use
-----------

For each tool you typically start by entering information about the list/library that you want to interact with.

1. Enter the display name or GUID of the list/library that you want to interact with
2. Enter the display name of the SharePoint site that you want to interact with.
3. Enter a CAML query to select the items from the list/library that you want to interact with.

From that point each tool differs depending on what actions it performs. There is also a **Preview** button that will run your CAML query to let you verify which list/library items you have selected. I recommend using this every time to make sure your CAML is correct.

### Update Items

This tool allows you to set the value of a field (or multiple fields) for multiple list items at once.

It currently only supports one action--Remove Lookup Item. This allows you to remove a single lookup from a field that allows multiple lookups.

1. Once your List Name, Web URL, and CAML Query fields are set, click the **Add a Field** button.
2. Select the multiple lookup field from the **Field** dropdown.
3. Select "Remove Lookup Item" from the **Update** dropdown.
4. Enter the lookup item to remove. Use the item ID, a semicolon, a hash symbol, and the display name of the lookup as used in the list (e.g. if the Title field is displayed in the lookup column, use the value of the title field after the ";#"). I plan on creating a dropdown of available lookup items to choose from so you don't have to manually enter the value, but for now this serves my needs.
5. Click the **Update Items** button.

### Approve List Items

This tool allows you to approve multiple list items that require content approval at once.

There are no additional steps needed beyond entering the list information, the CAML query, and clicking the **Approve Items** button.

### Dup List Items

This tool allows you to create a duplicate of a list item from one list into another list. It won't retain SharePoint's internal metadata like Created, Created by, Modified, or Modified by, but otherwise it can duplicate any fields that you specify. And the best part? You can remap fields from the source list to differently-named fields (of the same type) in the destination list.

1. Once your Source List Name, Source Web URL, CAML Query, Destination List Name, and Destination Web URL fields are set, click the **Add a Field** button.
2. Select the first field you want to copy in the **Source Field** dropdown.
3. Select the field from the **Destination Field** dropdown that you want to copy the source field into. If the destination list has a field with the same display name as the source field, it will auto-select that field for you.
4. Repeat this process for each field that you want to copy.
5. Click the **Duplicate Items** button.

### Check-in Documents

This tool allows you to check-in multiple files in a document library at once. Note that it requires a Server URL (i.e. the path to the site collection root) in addition to the standard list/library fields.

1. Once your Library Name, Server URL, Web URL, and CAML Query fields are set, select the type of check-in to perform.
2. Click the **Check in Documents** button.

### Start Workflow

This tool allows you to start a workflow on multiple list/library items at once.

1. Once your List Name, Web URL, and CAML Query are set, enter the display name of the Workflow.
2. Click the **Start Workflows** button.

### Check Permissions

This is still in development.

### Script Audit

This is still in development.

### Select Items

This tool allows you to output a list of items on the page based on your CAML Query and some additional filters. I created it as a quick and easy way to randomly select an item from a list, so that's the only "cool" function for the tool at the moment.

1. Once your List Name, Web URL, and CAML Query fields are set, select the additional filter you wish to perform when selecting the list items (if any).
2. Click the **View Results** button.
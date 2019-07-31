messages = new Array();

//clientview.asp
messages["logo_url"] = "Logos need to be uploaded to the server manually. Please email your company logo to your contact at Genergy.";

//fixtype.asp
messages["ballast_life_in_hours"] = "From manufacturer data.";
messages["ballast_type"] = "Manufacturer catalog number.";
messages["catalog_number"] = "Manufacturer's fixture model number";
messages["estimated_lamp_life_in_hours"] = "From manufacturer data - may require an adjustment factor based on occupancy and other factors.";
messages["fixture_code"] = "Floor- and fixture-specific tag used for keying to fixture schedule (unique to client location).";
messages["fixture_description"] = "Describe fixture size, lamp type (i.e. fluorescent, incandescent), mounting description (i.e. troffer, sconce, downlight, strip, etc), and any other special information.";
messages["general_remarks"] = "Miscellaneous information regarding special trim, unusual conditions, etc.";
messages["lamp_catalog_number"] = "ANSI standard code for particular lamp.";
messages["lamp_power"] = "Can be found in ANSI code or manufacturer data, in Watts.";
messages["manufacturer"] = "Refer to fixture nameplate.";
messages["voltage"] = "From manufacturer data.";

//mapsetup.asp
messages["new_map"] = "Enter the path to the map file. Maps need to be uploaded to the server manually. ";
messages["primary_map"] = "The primary map appears on the home page of the Lighting &amp; Maintenance application when you log in.";

//mapsetupedit.asp
messages["alt"] = "Alt text, also known as a tooltip, is the text that appears when you hover the mouse over the point on the map";
messages["link_map_pt_to"] = "Points on the map can be set to link to another map or to a node. If, for instance, you manage three buildings--two in Chicago,  one in Los Angeles--you might want to configure a map of the United States with two points: one linking to a map of Chicago, the other linking to the building node for the single facility in Los Angeles.<br><br><b>Link Point to Map</b><br>If you haven't yet added the map to which you want to link, skip this pulldown and return to edit this point later. If you have already added the desired map, select it from the pulldown now.<br><br><b>Select Node</b><br>A point on the map can expand the menu tree to show a node (region, service, city or building) that is nested deep in the tree structure. Click the 'Select Node' button to display your menu and pick a node to link.";
messages["pickmap"] = "Select map to edit";
messages["save_point"] = "New points will not be saved unless you click <b>Add Point</b>. Edits to existing points will likewise not be saved unless you click <b>Edit Point</b> when you are finished making your changes.";
messages["xy"] = "Manually alter the selected point's coordinates to finetune its position on the map";

//newfixture.asp
messages["new_fixture_in_rm"] = "The Fixture Type pulldown shows the fixtures that were created in the Fixture Management section. To add a fixture to this menu please create a new fixture in the Fixture Management section for this building.";

//treesetup.asp
messages["add_bldg_to_tree"] = "Buildings need to be created in the Facilities Manager before they can be added to the node tree. If there are no items in this pulldown, click the second button above, right, to go to the Facilities Manager and input a building.";
messages["click_node_first"] = "Pick node first, select action later! You cannot perform any action on the tree without first selecting a node. New items are added to the tree beneath the node you select.";
messages["link"] = "A node can link to a page or a map. To link to a map, you will first need to set up a map by clicking the Map Configure button in the upper right corner of this page. Then return here to select a map from the pulldown.";
messages["node_type"] = "Certain types of node have predefined label values. When you select Building or Service, you will be presented with a list from which to choose the label for your node. Region and City have no predefined values; selecting either will present you with the opportunity to type in a label.<br><br><b>Region</b> or <b>City</b><br>Organize your facilities by geographical location to make them easier to find in the expanding  menu<br><br><b>Building</b><br>Set the name of an individual node to match a building you have entered through the Facilities Manager. Please note that this option will not create any new nodes, it will simply reassign an existing node and its children to a building. To add a building with the standard set of services, return to step 2 and select 'Add building' from the 'Select an action' pulldown.<br><br><b>Service</b><br>Assign the selected node a standard service label or select &quot;Add new service&quot; from the top of the label pulldown. Any labels that you type in will be added to the  service pulldown for future use";
messages["select_tree_action"] = "<b>Add building</b><br>Adds a standard set of services for a building entered in the Facilities Manager below the node you have selected<br><br><b>Add new node</b><br>Adds a single node to the tree which you can define as a region, city, address, or service<br><br><b>Edit/delete node</b><br>Allows you to change the label for the node you have selected<br><br><b>Move node</b><br>Lets you move the selected node beneath a different node (note that you can't move a node under one of its own children)<br><br><b>Copy node</b><br>Lets you duplicate a node";
messages["whats_a_label"] = "Name your node. The options that appear in the label pulldown are based on your selection of Node Type";

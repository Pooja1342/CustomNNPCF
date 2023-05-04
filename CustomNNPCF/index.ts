import {IInputs, IOutputs} from "./generated/ManifestTypes";
import * as ReactDOM from 'react-dom';
import * as React from 'react';
import swal from 'sweetalert2';
import ToggleComponent from "./components/ToggleComponent";
import { type } from "os";
import { getuid } from "process";

class detailItem {
	parent: string;
	key: string;
	value: string;
}

class selState {
	label: string;
	text: string;
	total: number;
	actual: number;
}

export class CustomNNPCF implements ComponentFramework.StandardControl<IInputs, IOutputs> {

   // Global Variables
	private _context: ComponentFramework.Context<IInputs>;
	private _container: HTMLDivElement;
	private _mainContainer: HTMLDivElement;
	private _unorderedList: HTMLUListElement;
	private _errorLabel: HTMLLabelElement;
	public _defaultFilter: string;
    public _expandFilter: string;
	public _filter: string;
	private _entityName: string;
	private _selectorLabel: string;
    private _selectorOrder : string;
	public _selValues: detailItem[];
	public _values: string[];
	private _checkBoxChanged: (evnt: Event) => void;
	private _notifyOutputChanged: () => void;
	private _togglePanel: HTMLDivElement;
	private _itemList: detailItem[];
	private _selStates: selState[];
	private _showToggle = false;

	constructor()
	{

	}

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public async init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
	{
		this._context = context;
		this._container = container;
		this._mainContainer = document.createElement("div");
		this._unorderedList = document.createElement("ul");
		this._errorLabel = document.createElement("label");
		this._unorderedList.classList.add("ks-cboxtags");
		this._mainContainer.classList.add("multiselect-container");
		this._itemList = [];

		if (this._context.parameters.visibilityToggle.raw != null) {
			this._showToggle = this._context.parameters.visibilityToggle.raw == "1" ? true : false;
		}

		if (this._context.parameters.defaultFilter.raw != null) {
			this._defaultFilter = this._context.parameters.defaultFilter.raw;
		}

        if (this._context.parameters.expandFilter.raw != null) {
			this._expandFilter = this._context.parameters.expandFilter.raw;
		}


		//Trigger function on check-box change.
		this._notifyOutputChanged = notifyOutputChanged;
		this._checkBoxChanged = this.checkBoxChanged.bind(this);
		this._selStates = [];

  
	/* 	if (Xrm.Page.ui.getFormType() !== 1) {
			await this.getRelatedRecords();
		}
 */
        // @ts-ignore
		var contextInfo = this._context.mode.contextInfo;
        var recordId = contextInfo.entityId;
        if(recordId != null){
			await this.getRelatedRecords();
		}
	}


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public async updateView(context: ComponentFramework.Context<IInputs>)
    {
        // Add code to update control view
        this._context = context;

		//Check that the entityId value has been updated before refreshing the control
		if (context.updatedProperties != null && context.updatedProperties.length != 0) {
			if (context.updatedProperties[context.updatedProperties.length - 1] == "entityId" || context.updatedProperties[context.updatedProperties.length - 1] == "IsControlDisabled") {
				await this.getRecords();
			}
		}
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void
    {
        // Add code to cleanup control if necessary
    }

    public async getRelatedRecords() {
        try {
            var list = [];
            // @ts-ignore
            var contextInfo = this._context.mode.contextInfo;
            var recordId = contextInfo.entityId;
            var lookupTo = this._context.parameters.lookuptoAssociatedTable.raw!.toLowerCase();
            var lookupFrom = this._context.parameters.lookuptoCurrentTable.raw!.toLowerCase();
            var result = await this._context.webAPI.retrieveMultipleRecords(this._context.parameters.associationTable.raw!, '?$select= _' + lookupTo + '_value&$filter=_' + lookupFrom + '_value eq ' + recordId);
            for (var i = 0; i < result.entities.length; i++) {
                var temp = result.entities[i]["_" + lookupTo + "_value"];
                list.push(temp);
                var cItem = this._itemList.find((e => e.key === temp));
                var cState = this._selStates.findIndex(e => e.text === cItem?.parent);
                if (cState !== -1) {
                    this._selStates[cState].actual++;
                }
            }
            this._values = list;
            await this.getRecords();
            } 
            catch(error:any)
             { 
                swal.fire("getRelatedRecords", "Error:" + error.message, "error");
            }
        }

        //Called to retrieve records to display, both on-load and on-change of lookup
        public async getRecords() {
            try {
                this._container.innerHTML = "";
                this._mainContainer.innerHTML = "";
                this._unorderedList.innerHTML = "";
                if (this._showToggle) {
                    this._togglePanel = document.createElement("div");
                    this._togglePanel.style.float = "right";
                    var toggleProps = {
                        visible: true,
                        onChangeResult: this.showHideControl.bind(this)
                    }
        
                    ReactDOM.render(React.createElement(ToggleComponent, toggleProps), this._togglePanel);
                    this._mainContainer.appendChild(this._togglePanel);
                }
                //Check if table name variable contains data
                if (this._context.parameters.selectorTable.raw != null && this._context.parameters.selectorTable.raw != "") {
                    this._entityName = this._context.parameters.selectorTable.raw;
                }
                //Check if field name contains data
                if (this._context.parameters.selectorLabel.raw != null && this._context.parameters.selectorLabel.raw != "") {
                    this._selectorLabel = this._context.parameters.selectorLabel.raw;
                }
                 //Check if any order by field contains data
                 if (this._context.parameters.selectorOrder.raw != null && this._context.parameters.selectorOrder.raw != "") {
                    this._selectorOrder = this._context.parameters.selectorOrder.raw;
                }
                this._filter = "?$select=" + this._selectorLabel + "," + this._entityName + "id" + "&$orderby=" + this._selectorOrder + " asc";
                //Check if both Default Filter and Expand Filter contains data
                if ((this._defaultFilter !== undefined && this._defaultFilter !== "") && (this._expandFilter !== undefined && this._expandFilter !== "")){
                    this._filter += "&$filter=" + this._defaultFilter + " and " + this._expandFilter;
                } 
                if((this._defaultFilter != undefined && this._defaultFilter != "") && (this._expandFilter == undefined || this._expandFilter == "")){
                    this._filter += "&$filter=" + this._defaultFilter;
                }
                if((this._defaultFilter == undefined || this._defaultFilter == "") && (this._expandFilter != undefined && this._expandFilter != "")) 
                {
                    this._filter += "&$filter=" + this._expandFilter;
                } 

                var records = await this._context.webAPI.retrieveMultipleRecords(this._entityName, this._filter);
                for (var i = 0; i < records.entities.length; i++) {
                    var newChkBox = document.createElement("input");
                    var newLabel = document.createElement("label");
                    var newUList = document.createElement("li");
        
                    newChkBox.type = "checkbox";
                    newChkBox.id = records.entities[i][this._entityName + "id"];
                    newChkBox.name = records.entities[i][this._selectorLabel];
                    newChkBox.value = records.entities[i][this._entityName + "id"];
                    if (this._values != undefined) {
                        if (this._values.includes(newChkBox.id)) {
                            newChkBox.checked = true;
                        }
                    }
                    newChkBox.addEventListener("change", this._checkBoxChanged);
                    newLabel.innerHTML = records.entities[i][this._selectorLabel];
                    newLabel.htmlFor = records.entities[i][this._entityName + "id"];
                    newUList.appendChild(newChkBox);
                    newUList.appendChild(newLabel);
                    this._unorderedList.appendChild(newUList);
                }
                this._mainContainer.appendChild(this._unorderedList);
                this._mainContainer.appendChild(this._errorLabel);
                this._container.appendChild(this._mainContainer);
        
                } catch(error:any) {
                    swal.fire("getRecords", "Error:" + error.message, "error");
                }
            }
            
            public async checkBoxChanged(evnt: Event) {
            try {
                var targetInput = <HTMLInputElement>evnt.target;
                // @ts-ignore
                var contextInfo = this._context.mode.contextInfo;
                var recordId = contextInfo.entityId;
                var thisEntity = contextInfo.entityTypeName;
                var thatEntity = this.getEntityPluralName(this._entityName);
                var thisEntityPlural = this.getEntityPluralName(thisEntity);
                var associationTable = this._context.parameters.associationTable.raw!;
                var lookupFieldTo = this._context.parameters.lookuptoAssociatedTable.raw!;
                var lookupFieldFrom = this._context.parameters.lookuptoCurrentTable.raw!;
                var lookupToLower = lookupFieldTo.toLowerCase();
                var lookupFromLower = lookupFieldFrom.toLowerCase();
                var lookupDataTo = lookupFieldTo + "@odata.bind";
                var lookupDataFrom = lookupFieldFrom + "@odata.bind";
                var associationTableNameField = this._context.parameters.associationLable.raw!;
        
                var data =
                {
                    [associationTableNameField]: targetInput.name,
                    [lookupDataTo]: "/" + thatEntity + "(" + targetInput.id + ")",
                    [lookupDataFrom]: "/" + thisEntityPlural + "(" + recordId + ")"
                }
                var actual = 0;
                var cState = this._selStates.findIndex(e => e.text === targetInput.value);
                if (cState !== -1)
                    actual = this._selStates[cState].actual;
        
                if (targetInput.checked) {
                    await this._context.webAPI.createRecord(associationTable, data);
                    actual++;
                }
                else {
                    await this.deleteRecord(associationTable, lookupToLower, targetInput.id, lookupFromLower, recordId);
                    actual--;
                }
            
                this._notifyOutputChanged();
                } catch (error:any) {
                    swal.fire("checkBoxChanged", "Error:" + error.message , "error");
                }
            }
        
            //Async delete record process called when a check-box is unchecked
            private async deleteRecord(associationTable: string, lookupToLower: string, targetInput: string, lookupFromLower: string, recordId: string) {
                let _this = this;
                try {
                    var result = await this._context.webAPI.retrieveMultipleRecords(associationTable, '?$select=' + associationTable + 'id&$filter=_' + lookupToLower + '_value eq ' + targetInput + ' and _' + lookupFromLower + '_value eq ' + recordId)
                    for (var i = 0; i < result.entities.length; i++) {
                        var linkRecordId = result.entities[i][associationTable + 'id'];
                    }
                    _this._context.webAPI.deleteRecord(associationTable, linkRecordId)
                } catch(error:any) { 
                    swal.fire("deleteRecord", "Error:" + error.message, "error");
                }
            }
        
            public async showHideControl(show: boolean) {
            try {
                var display = "inline";
                if (show === false) {
                    display = "none";
                }
                this._unorderedList.style.display = display;
                } catch (error:any) {
                    swal.fire("showHideControl", "Error:" + error.message, "error");
                }
            }
            
            public async refreshItems() {
            try {
                await this.getRecords();
                return true;
                } catch (error:any) {
                    swal.fire("refreshItems", "Error:" + error.message, "error");
                }
            }
        
            //Retrieve plural name of a table
            private getEntityPluralName(entityName: string): string {
                if (entityName.endsWith("s"))
                    return entityName + "es";
                else if (entityName.endsWith("y"))
                    return entityName.slice(0, entityName.length - 1) + "ies";
                else
                    return entityName + "s";
            }
    
}

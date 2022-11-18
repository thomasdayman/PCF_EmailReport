import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as $ from "jquery";
import { send } from "process";
import { resolve } from "dns";

let popUpContent: HTMLDivElement;

export class EmailReport implements ComponentFramework.StandardControl<IInputs, IOutputs> {

	// Value of the field is stored and used inside the control 
	private _value: string | null;

	// PCF framework context, "Input Properties" containing the parameters, control metadata and interface functions.
	private _context: ComponentFramework.Context<IInputs>;

	// PCF framework delegate which will be assigned to this object which would be called whenever any update happens. 
	private _notifyOutputChanged: () => void;

	private loadingLabel: HTMLLabelElement;
	private loadingSpinner: HTMLDivElement;

	private alertDialogDiv: HTMLDivElement;
	private successAlertDialog: HTMLDivElement;
	private failureAlertDialog: HTMLDivElement;
	private _popUpService: ComponentFramework.FactoryApi.Popup.PopupService;
	private popUpContentDiv: HTMLDivElement;

	private reportGuid: string;
	private reportName: string;

	// Control's container
	private controlContainer: HTMLDivElement;

	private runReportButtonButton: HTMLButtonElement;

	/**
	 * Empty constructor.
	 */
	constructor() {

	}

	/**
	 * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
	 * Data-set values are not initialized here, use updateView.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
	 * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
	 * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
	 * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
	 */
	public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {

		interface PopupDev extends ComponentFramework.FactoryApi.Popup.Popup {
			popupStyle: object;
		}

		this._context = context;
		this._notifyOutputChanged = notifyOutputChanged;
		this.controlContainer = document.createElement("div");

		this.loadingSpinner = document.createElement("div");
		this.loadingSpinner.id = "loadingSpinner";

		this.loadingLabel = document.createElement("label");
		this.loadingLabel.id = "loaderLabel";

		this.successAlertDialog = document.createElement("div");
		this.successAlertDialog.id = "success";

		this.failureAlertDialog = document.createElement("div");
		this.failureAlertDialog.id = "failure";

		this.alertDialogDiv = document.createElement("div");
		this.alertDialogDiv.id = "mainAlertDiv"
		this.alertDialogDiv.appendChild(this.successAlertDialog);
		this.alertDialogDiv.appendChild(this.failureAlertDialog);

		this.popUpContentDiv = document.createElement('div');
		this.popUpContentDiv.id = "mainpopUpDiv";
		this.popUpContentDiv.appendChild(this.loadingLabel);

		this.runReportButtonButton = document.createElement("button");
		this.runReportButtonButton.innerHTML = "Send Email";
		this.runReportButtonButton.addEventListener("click", this.onSendEmailClick.bind(this));

		popUpContent = document.createElement('div');
		popUpContent.id = "popUp";
		popUpContent.style.backgroundColor = "white";
		popUpContent.prepend(this.alertDialogDiv);
		popUpContent.appendChild(this.loadingSpinner);
		popUpContent.appendChild(this.popUpContentDiv);

		let popUpOptions: PopupDev = {
			closeOnOutsideClick: false,
			content: popUpContent,
			name: 'mainPopup',
			type: 1,
			popupStyle: {}
		};

		this.controlContainer.appendChild(this.runReportButtonButton);
		container.appendChild(this.controlContainer);
		this._popUpService = context.factory.getPopupService();
		this._popUpService.createPopup(popUpOptions);
	}

	public showMainPopUp(): void {
		this._popUpService.openPopup('mainPopup');

		if (this._context.parameters.automaticallySendReport.raw == "true") {
			this.loadingLabel.innerHTML = "Sending Email...";
		}
		else {
			this.loadingLabel.innerHTML = "Opening Email...";
		}

	}
	public closeMainPopUp(): void {
		this._popUpService.closePopup('mainPopup');
	}

	private onSendEmailClick(event: Event): void {
		try {
			this.showMainPopUp();
			this.runReport(this);
		}
		catch (err) {
			this.showFailureAlert(err as any, this);
		}
	}

	private showFailureAlert(error: Error, Object: this): void {
		this.failureAlertDialog.innerHTML = error.message;

		$("#failure").fadeIn(0).delay(2000).fadeOut(400).promise().done(function () {
			Object.closeMainPopUp();
		});
	}

	private showSuccessAlert(success: any, Object: this): void {
		this.successAlertDialog.innerHTML = success;

		$("#success").fadeIn(0).delay(2000).fadeOut(400).promise().done(function () {
			Object.closeMainPopUp();
		});
	}

	private GetEntityPluralName = async (entityName: string): Promise <any> => {
        let entityplural;
		
		await this._context.utils.getEntityMetadata(entityName)
		.then(			
			(result) => {
				entityplural = result.EntitySetName;
		},
		(err) => {
			this.showFailureAlert(err, this);
			entityplural = err;
		});

		return entityplural;
	}

	private runReport(context: any): void {
		const promise = new Promise((resolve, reject) => {
			this.executeReport();
		});
		promise.then((res) => {
		});
		promise.catch((err) => {
			this.showFailureAlert(err, this);
		});
	}

	private executeReport() {
		try {
			this.getReportingSession(this);
		}
		catch (err) {
			this.showFailureAlert(err as any, this);
		}
	}

	private getReportingSession(Object: this) {
		try {
			var recordId = (this._context.mode as any).contextInfo.entityId;
			recordId = recordId.replace('{', '').replace('}', '');

			var fetchXML = this._context.parameters.Report_FetchXML.raw?.toString().replace('{0}', this._context.parameters.RecordId.raw!);

			var pth = (<any>this._context).page.getClientUrl() + "/CRMReports/rsviewer/reportviewer.aspx";

			var request = new XMLHttpRequest();

			return new Promise(function (resolve, reject) {

				request.onreadystatechange = function () {

					if (request.readyState !== 4) return;

					if (request.status >= 200 && request.status < 300) {
						var x = request.responseText.lastIndexOf("ReportSession=");
						var y = request.responseText.lastIndexOf("ControlID=");

						var ret = new Array();
						ret[0] = request.responseText.substr(x + 14, 24);
						ret[1] = request.responseText.substr(y + 10, 32);

						Object.encodePdf(ret, Object);
					} else {
						Object.showFailureAlert(Error(request.statusText), Object);
					}
				};

				request.open('POST', pth, true);
				request.setRequestHeader("Accept", "*/*");
				request.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

				request.send("id=%7B" + Object._context.parameters.reportGuid.raw + "%7D&uniquename=" + (<any>Object._context).orgSettings.uniqueName + "&iscustomreport=true&reportnameonsrs=&reportName=" + Object._context.parameters.reportName.raw + "&isScheduledReport=false&p:" + Object._context.parameters.Report_Parameter_Name.raw + "=" + fetchXML);
			});
		}
		catch (err) {
			this.showFailureAlert(err as any, this);
		}
	}

	private encodePdf(responseSession: any, Object: this) {
		try {
			var retrieveEntityReq = new XMLHttpRequest();
			var pth = (<any>this._context).page.getClientUrl() + "/Reserved.ReportViewerWebControl.axd?ReportSession=" + responseSession[0] + "&Culture=1033&CultureOverrides=True&UICulture=1033&UICultureOverrides=True&ReportStack=1&ControlID=" + responseSession[1] + "&OpType=Export&FileName=Public&ContentDisposition=OnlyHtmlInline&Format=PDF";
			retrieveEntityReq.open("GET", pth, true);
			retrieveEntityReq.setRequestHeader("Accept", "*/*");
			retrieveEntityReq.responseType = "arraybuffer";
			retrieveEntityReq.onreadystatechange = function () {

				if (retrieveEntityReq.readyState == 4 && retrieveEntityReq.status == 200) {
					var binary = "";
					var bytes = new Uint8Array(this.response);

					for (var i = 0; i < bytes.byteLength; i++) {
						binary += String.fromCharCode(bytes[i]);
					}
					var bdy = btoa(binary);
					Object.createEmail(bdy, Object);
				}
			};
			retrieveEntityReq.send();
		}
		catch (err) {
			this.showFailureAlert(err as any, this);
		}
	}

	private async createEmail(data: any, Object: this) {
		try {
			var recordId = (this._context.mode as any).contextInfo.entityId;
			recordId = recordId.replace('{', '').replace('}', '');

			let entity: any = {};

			if(this._context.parameters.subject.raw !== null)
			{
				entity["subject"] = this._context.parameters.subject.raw;
			}

			if(this._context.parameters.description.raw !== null)
			{
				entity["description"] = this._context.parameters.description.raw;
			}

			if(this._context.parameters.RegardingObjectEntityName.raw !== null || this._context.parameters.RegardingObjectId.raw)
			{
				entity["regardingobjectid_" + this._context.parameters.RegardingObjectEntityName.raw + "@odata.bind"] = "/" + await this.GetEntityPluralName(this._context.parameters.RegardingObjectEntityName.raw!) + "(" + this._context.parameters.RegardingObjectId.raw + ")";
			}

			let activityparties = [];

			let from: any = {};
			if(this._context.parameters.From.raw !== null)
			{
				from["partyid_systemuser@odata.bind"] = "/" + this._context.parameters.FromEntityType.raw + "s(" + this._context.parameters.From.raw + ")";
				from["participationtypemask"] = 1;
				activityparties.push(from);
			}

			let to: any = {};
			if(this._context.parameters.To.raw !== null)
			{
				to["partyid_contact@odata.bind"] = "/" + this._context.parameters.ToEntityType.raw + "s(" + this._context.parameters.To.raw + ")";
				to["participationtypemask"] = 2;
				activityparties.push(to);
			}

			//activityparties.push(to);
			//activityparties.push(from);
			if(activityparties.length != 0)
			{
				entity["email_activity_parties"] = activityparties;
			}

			this._context.webAPI.createRecord("email", entity).then(
				(result) => {
					var newEntityId = result.id;
					this.createEmailAttachment(data, newEntityId.toString(), Object);
				},
				(err) => {
					Object.showFailureAlert(err, Object);
				}
			);
		}
		catch (err) {
			this.showFailureAlert(err as any, this);
		}
	}

	private createEmailAttachment(data: any, EmailId: string, Object: this) {
		try {
			var recordId = (this._context.mode as any).contextInfo.entityId;
			recordId = recordId.replace('{', '').replace('}', '');

			let entity: any = {};
			entity["objectid_activitypointer@odata.bind"] = "/activitypointers(" + EmailId + ")";
			entity.subject = "File Attachment";
			entity.body = data;
			entity.filename = this._context.parameters.attachmentFileName.raw;
			entity.objecttypecode = 'email';
			entity.mimetype = "application/pdf";

			this._context.webAPI.createRecord("activitymimeattachment", entity).then(
				(result) => {
					var newEntityId = result.id;

					if (this._context.parameters.automaticallySendReport.raw == "true") {
						this.sendEmail(EmailId, Object);
					}
					else {
						this.openEmailRecord(EmailId, Object);
					}
				},
				(err) => {
					this.showFailureAlert(err, this);
				}
			);
		}
		catch (err) {
			this.showFailureAlert(err as any, this);
		}
	}

	private openEmailRecord(newEntityId: string, Object: this) {
		try {
			let entityFormOptions: any = {};
			entityFormOptions.entityName = "email";
			entityFormOptions.entityId = newEntityId;

			// Open the form.
			this._context.navigation.openForm(entityFormOptions).then(
				function (success) {
					console.log(success);
				},
				function (err) {
					Object.showFailureAlert(err, Object);
				});
		}
		catch (err) {
			this.showFailureAlert(err as any, this);
		}
	}


	private sendEmail(EmailId: string, Object: this) {
		try {
			var parameters =
			{
				"IssueSend": true
			};

			var request = new XMLHttpRequest();
			request.open("POST", (<any>this._context).page.getClientUrl() + "/api/data/v9.1/emails(" + EmailId + ")/Microsoft.Dynamics.CRM.SendEmail", true);
			request.setRequestHeader("Accept", "application/json");
			request.setRequestHeader("Content-Type", "application/json; charset=utf-8");
			request.setRequestHeader("OData-MaxVersion", "4.0");
			request.setRequestHeader("OData-Version", "4.0");
			request.onreadystatechange = function () {
				if (this.readyState === 2) { //4
					request.onreadystatechange = null;
					if (this.status === 200) { //204
						Object.showSuccessAlert("Email has been sent!", Object);
					}
					else {
						Object.showFailureAlert(Error(this.statusText), Object);
					}
				}
				else {
					Object.showFailureAlert(Error(this.statusText), Object);
				}
			};
			request.send(JSON.stringify(parameters));
		}
		catch (err) {
			this.showFailureAlert(err as any, this);
		}
	}

	/**
	 * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
	 * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
	 */
	public updateView(context: ComponentFramework.Context<IInputs>): void {
		// Add code to update control view
	}

	/** 
	 * It is called by the framework prior to a control receiving new data. 
	 * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
	 */
	public getOutputs(): IOutputs {
		return {};
	}

	/** 
	 * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
	 * i.e. cancelling any pending remote calls, removing listeners, etc.
	 */
	public destroy(): void {
		// Add code to cleanup control if necessary
	}
}

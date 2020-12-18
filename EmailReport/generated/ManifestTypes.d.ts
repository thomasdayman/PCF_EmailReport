/*
*This is auto generated from the ControlManifest.Input.xml file
*/

// Define IInputs and IOutputs Type. They should match with ControlManifest.
export interface IInputs {
    sampleProperty: ComponentFramework.PropertyTypes.StringProperty;
    subject: ComponentFramework.PropertyTypes.StringProperty;
    description: ComponentFramework.PropertyTypes.StringProperty;
    reportName: ComponentFramework.PropertyTypes.StringProperty;
    reportGuid: ComponentFramework.PropertyTypes.StringProperty;
    attachmentFileName: ComponentFramework.PropertyTypes.StringProperty;
    automaticallySendReport: ComponentFramework.PropertyTypes.EnumProperty<"true" | "false">;
    FromEntityType: ComponentFramework.PropertyTypes.EnumProperty<"systemuser" | "queue">;
    From: ComponentFramework.PropertyTypes.StringProperty;
    ToEntityType: ComponentFramework.PropertyTypes.EnumProperty<"contact" | "account" | "systemuser" | "entitlement" | "lead" | "queue" | "knowledgearticle">;
    To: ComponentFramework.PropertyTypes.StringProperty;
    Report_Parameter_Name: ComponentFramework.PropertyTypes.StringProperty;
    Report_FetchXML: ComponentFramework.PropertyTypes.StringProperty;
    RecordId: ComponentFramework.PropertyTypes.StringProperty;
    RegardingObjectId: ComponentFramework.PropertyTypes.StringProperty;
    RegardingObjectEntityName: ComponentFramework.PropertyTypes.StringProperty;
}
export interface IOutputs {
    sampleProperty?: string;
}

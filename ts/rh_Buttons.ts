import { anyFieldDirty, ViewFormSection, lockFieldsExcept, gridIsDirty } from "./shared/common"; // This imports functions from the shared common file
import { AccountRelationshipTypes, StateCode, AccountStatusCode } from "./shared/optionsets"; // This imports an optionset from the shared optionsets file
import { SecurityRoles, userHasRoles } from "./shared/security"; // This imports a function from the shared security file
import { refreshTotalFields } from "./shared/totals";

export class Buttons {
    // NOTE: Important to grab security role in onLoad, but also doesn't always refresh as OnLoad not always fired before button rule runs
    public static userCommercialAdmin = false;
    public static statusReason: AccountStatusCode;

    public static async onLoad(executionContext: Xrm.Events.EventContext) {
        const formContext = executionContext.getFormContext();

        this.userCommercialAdmin = userHasRoles([SecurityRoles.CommercialAdministrator]);
        this.statusReason = formContext.getAttribute("statuscode").getValue();

        // Refresh ribbon
        formContext.ui.refreshRibbon();
    }

    // *** Button rules/commands ***

    // NOTE: Simple call to return whether user has commercial admin role
    public static async enableRuleSubmitButton() {
        return this.userCommercialAdmin;
    }

    // Example submit command
    public static async commandSubmitRecord(formContext: Xrm.FormContext) {
        // Check if editable subgrid has been saved first
        if (gridIsDirty(formContext.getControl("subgrid_example"))) {
            await Xrm.Navigation.openAlertDialog({
                text: "Please save the record in the grid before submitting.",
            });
        } else {
            // Remember to set status and state if changing state
            formContext.getAttribute("statuscode").setValue(AccountStatusCode.Inactive);
            formContext.getAttribute("statecode").setValue(StateCode.Inactive);

            try {
                await formContext.data.save();

                // Force status change in memory to prevent grid re-editing after submission
                this.statusReason = AccountStatusCode.Inactive;

                // tslint:disable-next-line: max-line-length
                await Xrm.Navigation.openAlertDialog({
                    text: `Your record has been submitted.`,
                });
            } catch (error: any) {
                Xrm.Navigation.openAlertDialog({ text: `Submission failed: ${error.message}` });
            }
        }
    }

    // Example cancel command
    public static async commandCancelRecord(formContext: Xrm.FormContext) {
        formContext.getAttribute("statuscode").setValue(AccountStatusCode.Cancelled);
        formContext.getAttribute("statecode").setValue(StateCode.Inactive);

        try {
            await formContext.data.save();

            // Force status change in memory to prevent grid re-editing after cancellation
            this.statusReason = AccountStatusCode.Cancelled;

            // tslint:disable-next-line: max-line-length
            await Xrm.Navigation.openAlertDialog({
                text: `Record has been cancelled.`,
            });
        } catch (error: any) {
            Xrm.Navigation.openAlertDialog({ text: `Cancelling record failed: ${error.message}` });
        }
    }


    // Open a form/navigate to
    // NOTE: This code opens a case specifying a specific BPF process
    // The process flow is specified as text in the rh_processflow field
    // This is the only way to specify opening a new record with a different BPF to the default
    public static async commandNewCase(formContext: Xrm.FormContext) {
        const sourceRef = formContext.data.entity.getEntityReference(); // Source entity ref
        const processFlowAttrib: Xrm.Attributes.StringAttribute = formContext.getAttribute("rh_processflow");
        const processFlow = processFlowAttrib.getValue() ?? "";

        const pageInput: Xrm.Navigation.PageInputEntityRecord = {
            createFromEntity: sourceRef,
            entityName: "incident",
            pageType: "entityrecord",
        };

        if (processFlow.length > 0) {
            // Retrieve process flow with same name
            try {
                const ret = await Xrm.WebApi.retrieveMultipleRecords(
                    "workflow",
                    `?$top=2&$select=workflowid&$filter=name eq '${encodeURI(processFlow)}' and statecode eq 1`,
                );

                if (ret?.entities?.length !== 1) {
                    return Xrm.Navigation.openAlertDialog({
                        text: `${ret?.entities?.length ?? 0
                            } Business Process Flows (BPFs) found with name '${processFlow}'. Please review the 'Process Flow' field on the *** source ref record type ***.`,
                        title: "Cannot create new case",
                    });
                }

                // Set BPF
                pageInput.processId = ret.entities[0].workflowid;
            } catch (error: any) {
                return Xrm.Navigation.openErrorDialog({
                    details: `Error opening new case form`,
                    message: error.message,
                });
            }
        }

        return Xrm.Navigation.navigateTo(pageInput);
    }
}
import { anyFieldDirty, gridIsDirty, lockFieldsExcept, setDecisionByOn, ViewFormSection } from "./shared/common"; // This imports functions from the shared common file
import { AccountRelationshipTypes, StateCode, AccountStatusCode } from "./shared/optionsets"; // This imports an optionset from the shared optionsets file
import { userHasRoles } from "./shared/security"; // This imports a function from the shared security file
import { refreshTotalFields } from "./shared/totals";

// This is defined for lookup filtering later
const clientFilter = `<filter type="and"><condition attribute='customertypecode' operator='eq' value='3' /><condition attribute='statecode' operator='eq' value='0' /></filter>`;

const TOTAL_FIELDS: string[] = ["dxc_actualcost", "dxc_totalcost"];

export class General {
    public static statusReason: AccountStatusCode;

    // *** Events *** //
    public static async onLoad(executionContext: Xrm.Events.EventContext) {
        // NOTE: If any calls in a method are async, then the function itself must be async, and this is contagious, i.e. it must go all the way up the call stack

        const formContext = executionContext.getFormContext();

        // Set status into global variable for other functions
        this.statusReason = formContext.getAttribute("statuscode").getValue();

        // Call code depending on form type
        if (formContext.ui.getFormType() === XrmEnum.FormType.Create || formContext.ui.getFormType() === XrmEnum.FormType.Update) {
            // Do some work here
        }

        // This function will remove certain values from an optionset list. They must be defined in optionsets.ts
        this.removeOptionsetValues(executionContext);

        // This function will change tab and section visibility
        this.setFieldsBasedOnRelationshipType(executionContext);

        // This code checks if an optionset has data or not
        const relTypeAttrib: Xrm.Attributes.OptionSetAttribute = formContext.getAttribute("customertypecode");
        if (!relTypeAttrib.getValue()) {
            // Run some code here if value is not set
        }

        // This function will set the type of a multi-entity type lookup
        this.setMultiLookupType(executionContext);

        // Save the record (async)
        await formContext.data.save();
    }

    public static async onChangeRelationshipType(executionContext: Xrm.Events.EventContext) {
        // NOTE: This sample gets the current record id, and uses it on a call to a function to check if
        // any contacts exist for this account
        const formContext = executionContext.getFormContext();
        const accountid = formContext.data.entity.getId(); // This gets current record id

        if (await this.multipleRecordRetrieve(accountid)) {
            await Xrm.Navigation.openErrorDialog({
                message: `Contacts already exist for this account`,
            });
        } else {
            // Logic for if no contacts exist
        }
    }

    public static onSave(eventContext: Xrm.Events.SaveEventContext) {
        // NOTE: Useful if custom totalling functionality has been setup, perhaps on a parent entity
        // An IsBatch field can be used to prevent recalculation everytime one of hte dependent records is altered, until
        // A retotalling operation is finally called

        const formContext = eventContext.getFormContext();
        const isbatchAttrib: Xrm.Attributes.BooleanAttribute = formContext.getAttribute("dxc_isbatch");
        if (isbatchAttrib && isbatchAttrib.getValue() === true) {
            isbatchAttrib.setValue(false);
            isbatchAttrib.setSubmitMode("always");
        }

        // Get save mode
        const saveMode = eventContext.getEventArgs().getSaveMode();

        if (saveMode === XrmEnum.SaveMode.Deactivate) {
            this.disableFields(eventContext.getFormContext(), true);
        }
    }

    // NOTE: this function disables required fields if deactivating record
    private static disableFields(formContext: Xrm.FormContext, isDisabled: boolean) {
        const nameCtrl: Xrm.Controls.StringControl = formContext.getControl("dxc_name");

        nameCtrl.setDisabled(isDisabled);
    }

    // *** Dialogs *** //
    private static async confirmDialog(executionContext: Xrm.Events.EventContext) {
        // NOTE: We are waiting on the return value from a dialog, so this method must be async

        const ret = await Xrm.Navigation.openConfirmDialog(
            {
                confirmButtonLabel: "OK",
                text: `Input your message to confirm here. Do you want to continue?`,
                title: "Dialog box title",
            },
            {
                height: 180,
                width: 550,
            },
        );

        if (!ret.confirmed) {
            await Xrm.Navigation.openErrorDialog({ message: `Enter your not confirmed message here.` });
            // NOTE: Carry out any code changes you need to on "not" confirm here, i.e. Cancel was clicked
            // Equally this if/else could be ret.confirmed for if OK is clicked
            // The next return will drop the function out and prevent any further processing
            return;
        }
    }

    private static async showProgressDialog(formContext: Xrm.FormContext) {
        // NOTE: Handle in a try/catch with a finally so if the async action you're running fails, we clear the progress indicator
        try {
            Xrm.Utility.showProgressIndicator("Please wait...");
            // Do something with an await

            await formContext.data.refresh(false);
        } catch (error: any) {
            await Xrm.Navigation.openAlertDialog({
                title: "Unable to do stuff",
                text: error.message,
            });
            return;
        } finally {
            Xrm.Utility.closeProgressIndicator();

            formContext.data.refresh(false);
        }
    }

    // *** Manipulate fields *** //
    private static async setFieldExamples(formContext: Xrm.FormContext) {
        // NOTE: Examples of setting different field types
        const dateAttrib: Xrm.Attributes.DateAttribute = formContext.getAttribute("dxc_datefield");
        dateAttrib.setValue(new Date());
    }

    private static removeOptionsetValues(executionContext: Xrm.Events.EventContext) {
        const formContext = executionContext.getFormContext();

        // Unless user is a system administrator restrict "Relationship Type" when adding an account to:
        // "Prospect", "Property Account", "Other" or "Influencer".
        if (formContext.ui.getFormType() === XrmEnum.FormType.Create && !userHasRoles(["System Administrator"])) {
            const relationshipTypeCtrl: Xrm.Controls.OptionSetControl = formContext.getControl("customertypecode");
            relationshipTypeCtrl.removeOption(AccountRelationshipTypes.Client);
            relationshipTypeCtrl.removeOption(AccountRelationshipTypes.Supplier);
        }
    }

    private static setFieldsBasedOnRelationshipType(executionContext: Xrm.Events.EventContext) {
        const formContext = executionContext.getFormContext();

        const frameworkTab = formContext.ui.tabs.get("tab_framework"); // Reference using tab name

        const relationshipType = formContext.getAttribute("customertypecode").getValue();

        if (relationshipType === AccountRelationshipTypes.Client) {
            /* This is a property account */
            frameworkTab.setVisible(false); // Make tab hidden

            ViewFormSection(formContext, "section_geocoding", false); // This finds a section by name and hides it
        } else {
            frameworkTab.setVisible(true); // Make tab visible

            ViewFormSection(formContext, "section_geocoding", true); // This finds a section by name and makes it visible
        }
    }

    private static setMultiLookupType(executionContext: Xrm.Events.EventContext) {
        const formContext = executionContext.getFormContext();

        // NOTE: This is an enforced example, and wouldn't work on the account form. It needs a lookup that is multi-type
        // This example being parentcustomerid off the contact entity, which can be account or contact
        // It is now possible to create these lookups
        const accountLookup: Xrm.Controls.LookupControl = formContext.getControl("parentcustomerid");
        const account = accountLookup.getAttribute().getValue();

        accountLookup.setEntityTypes(["account"]);

        // If account is set
        if (account) {
            // Some more code here
        }
    }

    private static async setFieldRequirement(formContext: Xrm.FormContext) {
        const customLookupControl: Xrm.Controls.LookupControl = formContext.getControl("dxc_customlookupid");

        // Set to required
        customLookupControl.getAttribute().setRequiredLevel("required");

        // Set to not required
        customLookupControl.getAttribute().setRequiredLevel("none");
    }

    private static async setFieldDisabled(formContext: Xrm.FormContext) {
        const customLookupControl: Xrm.Controls.LookupControl = formContext.getControl("dxc_customlookupid");

        // Disable (lock) field
        customLookupControl.setDisabled(true);

        // Enable field
        customLookupControl.setDisabled(false);
    }

    private static async stringManipulation(executionContext: Xrm.Events.EventContext) {
        const formContext = executionContext.getFormContext();
        const mobileAttrib: Xrm.Attributes.StringAttribute = formContext.getAttribute("mobilephone");
        const numberstr = mobileAttrib.getValue() ?? "";

        if (numberstr.includes("+44") || numberstr.includes(" ")) {
            formContext.ui.setFormNotification("Please specify mobile number in format 07xxxxxxxxx with no spaces or +44", "ERROR", "mobilenumber");
        } else {
            formContext.ui.clearFormNotification("mobilenumber");
        }

        if (numberstr.length !== 11) {
            formContext.ui.setFormNotification(
                "Mobile number should be 11 digits. Please specify mobile number in format 07xxxxxxxxx with no spaces or +44",
                "ERROR",
                "mobilelength",
            );
        } else {
            formContext.ui.clearFormNotification("mobilelength");
        }
    }

    private static async setCustomLookupFilter(formContext: Xrm.FormContext) {
        // NOTE: Call this from another function and pass in formcontext
        // This will apply a custom fetch XML filter (defined in a const at the start of this class) to a lookup

        try {
            const customerLookup: Xrm.Controls.LookupControl = formContext.getControl("customerid");

            customerLookup.setEntityTypes(["account"]);
            customerLookup.addPreSearch(this.filterClientAccounts);
        } catch (err) {
            Xrm.Navigation.openAlertDialog({
                text: `Error ${err} setting filters on client lookup`,
            });
        }
    }

    private static filterClientAccounts(executionContext: Xrm.Events.EventContext) {
        const formContext = executionContext.getFormContext();
        const accountLookup: Xrm.Controls.LookupControl = formContext.getControl("customerid");
        accountLookup.addCustomFilter(clientFilter, "account");
    }

    private static applyCustomXml(formContext: Xrm.FormContext) {
        // NOTE: This complex example applies custom XML (criteria and columns) to a lookup on the fly
        const lookupCtrl: Xrm.Controls.LookupControl = formContext.getControl("dxc_lookupid");

        const workstreamId = "00000000-0000-0000-0000-000000000000"; // Enforced example

        const today = new Date().toISOString();
        let tenancyType: number = 1; // NOTE: Enforced example, you would retrieve this from somewhere
        let propertyAccountType: number = 1; // NOTE: Enforced example, you would retrieve this from somewhere

        const fetchXml = [
            "<fetch>",
            "  <entity name='ebecs_repairstandard'>",
            "    <attribute name='ebecs_name' />",
            "    <attribute name='ebecs_repairtypeid' />",
            "    <attribute name='ebecs_repairagreementid' />",
            "    <attribute name='ebecs_repairpriority' />",
            "    <attribute name='ebecs_chargeability' />",
            "    <attribute name='ebecs_repairresponsibility' />",
            "    <filter>",
            "      <condition attribute='statecode' operator='eq' value='0' />",
            "    </filter>",
            "    <link-entity name='ebecs_repairtype' from='ebecs_repairtypeid' to='ebecs_repairtypeid' link-type='inner' alias='rt'>",
            "      <attribute name='ebecs_description' />",
            "    </link-entity>",
            "    <link-entity name='ebecs_repairagreement' from='ebecs_repairagreementid' to='ebecs_repairagreementid' link-type='inner' alias='ra'>",
            "      <attribute name='ebecs_applicablepropertytypes' />",
            "      <attribute name='ebecs_applicabletenancytypes' />",
            "      <filter>",
            "        <condition attribute='statuscode' operator='eq' value='1' />",
            `        <condition attribute='ebecs_startdate' operator='le' value='${today}' />`,
            `        <condition attribute='ebecs_enddate' operator='ge' value='${today}' />`,
            "        <filter type='or'>",
            "          <condition attribute='ebecs_isselectbypropertytype' operator='ne' value='1' />",
            //         Property (account) property type
            "          <condition attribute='ebecs_applicablepropertytypes' operator='contain-values'>",
            `            <value>${propertyAccountType}</value>`,
            "          </condition>",
            "        </filter>",
            "        <filter type='or'>",
            "          <condition attribute='ebecs_isselectbytenancytype' operator='ne' value='1' />",
            //         Property (account) tenancy type
            "          <condition attribute='ebecs_applicabletenancytypes' operator='contain-values'>",
            `            <value>${tenancyType}</value>`,
            "          </condition>",
            "        </filter>",
            "      </filter>",
            "      <link-entity name='ebecs_repairagreementworkstream' from='ebecs_repairagreementid' to='ebecs_repairagreementid' link-type='inner' intersect='true'>",
            "        <link-entity name='ebecs_workstream' from='ebecs_workstreamid' to='ebecs_workstreamid' intersect='true'>",
            "          <filter>",
            `            <condition attribute='ebecs_workstreamid' operator='eq' value='${workstreamId}' />`,
            "          </filter>",
            "        </link-entity>",
            "      </link-entity>",
            "    </link-entity>",
            "  </entity>",
            "</fetch>",
        ]
            .map((x) => x.trim())
            .join("");

        const layoutXml = [
            "<grid name='ebecs_repairstandards' jump='ebecs_name' select='1' icon='1' preview='0'>",
            "  <row name='ebecs_repairstandard' id='ebecs_repairstandardid'>",
            "    <cell name='rt.ebecs_description' width='251' />",
            "    <cell name='ebecs_repairpriority' width='100' />",
            "    <cell name='ebecs_chargeability' width='100' />",
            "    <cell name='ebecs_repairresponsibility' width='100' />",
            "    <cell name='ra.ebecs_applicablepropertytypes' width='300' />",
            "    <cell name='ra.ebecs_applicabletenancytypes' width='300' />",
            "    <cell name='ebecs_name' width='100' />",
            "    <cell name='ebecs_repairagreementid' width='300' />",
            "  </row>",
            "</grid>",
        ]
            .map((x) => x.trim())
            .join("");

        lookupCtrl.addCustomView(
            "3a38d7db-a9e9-4da3-99f5-201dfeddd0f7", // Random GUID
            "dxc_entityname",
            "Title of custom view",
            fetchXml,
            layoutXml,
            true,
        );
    }

    private static async triggerBusinessRule(formContext: Xrm.FormContext) {
        // NOTE: When editing a field that has a business rule set on it, trigger the business rule
        const busRuleFieldAttrib: Xrm.Attributes.NumberAttribute = formContext.getAttribute("dxc_customnumberfield");
        busRuleFieldAttrib.setValue(500); // Set quantity
        busRuleFieldAttrib.fireOnChange(); // Trigger BR
    }

    private static async gridlineSelect(formContext: Xrm.FormContext) {
        // NOTE: THis locks all fields on an editable grid, except for the below fields
        // This method must be registered against the onLineSelect method of an editable grid
        // It must be included in the TS file against the parent entity the subgrid is on

        lockFieldsExcept(formContext, ["dxc_field1", "dxc_field2"]);

        // Couple this with security role/team checks to selectively enable fields on an editable grid based on role
    }

    // *** Retrieve records *** //
    private static async singleRecordRetrieve(executionContext: Xrm.Events.EventContext) {
        // NOTE: This example is taking an account id from a lookup, and if it's set, query a related record with that info

        const formContext = executionContext.getFormContext();

        // NOTE: Best to wrap in a try catch to grab any error returned
        try {
            const account = formContext.getAttribute("parentaccountid").getValue();
            if (!account) {
                return; // Drop out of processing as account not set
            }

            // NOTE: This example is expanding and retrieving a property off a related record
            const accountInfo = await Xrm.WebApi.retrieveRecord(
                "account",
                account[0].id,
                "?$select=name,customertypecode&$expand=ebecs_billingaccountid_Account($select=ebecs_prefix)",
            );
            const clientPrefix = accountInfo.ebecs_billingaccountid_Account?.ebecs_prefix;

            // NOTE: We can directly compare a returned field type to an optionset enum
            if (accountInfo.customertypecode == AccountRelationshipTypes.Client) {
                // More code here
            }

            if (clientPrefix) {
                // Use this in another operation
            }
        } catch (err) {
            Xrm.Navigation.openAlertDialog({ text: `Error ${err} obtaining Client prefix` });
        }
    }

    private static async multipleRecordRetrieve(accountId: string): Promise<boolean> {
        // NOTE: This function can be called from another to return a true/false depending if records exist
        // In this example, checking if any contacts exist with this accountid as parent account
        // It also uses StateCode enum from Optionsets shared file
        try {
            const ret = await Xrm.WebApi.retrieveMultipleRecords(
                "contact",
                "?$top=1" +
                    "&$select=parentaccountid" +
                    `&$filter=_parentaccountid_value eq ${accountId} and ` +
                    // tslint:disable-next-line: max-line-length
                    `Microsoft.Dynamics.CRM.In(PropertyName = 'statecode', PropertyValues = ['${StateCode.Active}'])`,
            );

            if (ret && ret.entities && ret.entities.length > 0) {
                // Return true if records found
                return true;
            }
        } catch (error: any) {
            await Xrm.Navigation.openErrorDialog({ message: `Error: ${error.message} checking for contacts` });
        }
        return false;
    }

    private static async asyncUpdateRecord(formContext: Xrm.FormContext) {
        // Calls async update record. Ideal for not updating the record you're on.
        await Xrm.WebApi.updateRecord("account", formContext.data.entity.getId(), {
            statuscode: AccountStatusCode.Inactive,
            statecode: StateCode.Inactive,
        });
    }

    // *** Miscellaneous *** //
    private static async preventFieldSaves(formContext: Xrm.FormContext) {
        const attrib: Xrm.Attributes.Attribute = formContext.getAttribute("dxc_customfield");
        // Not all fields are on all forms, so test if attribute is found.
        if (attrib) {
            attrib.setSubmitMode("never");
        }
    }

    private static async asyncFieldRefresh(formContext: Xrm.FormContext) {
        // NOTE: This code is for refreshing fields from the DB without refreshing the entire form
        // It has been used on totals fields that are refreshed in background, but other use cases exist
        // In this example, we check that the tab is expanded before refreshing

        const tab = formContext.ui.tabs.get("tab_summary");
        if (tab.getDisplayState() === "expanded") {
            await refreshTotalFields(formContext, TOTAL_FIELDS);
        }
    }

    private static async checkIfGridIsDirty(formContext: Xrm.FormContext) {
        // NOTE: This function checks if an editable subgrid is dirty, i.e. has not yet been saved
        if (gridIsDirty(formContext.getControl("subgridname"))) {
            await Xrm.Navigation.openAlertDialog({
                text: "Please save the record in the grid before submitting.",
            });
        }
    }

    private static async callBoundAction(formContext: Xrm.FormContext) {
        // eg Supplier application, fetchWorkstreamStatus
        const actionName = "ebecs_Approve";

        const dxcApproveRecord = {
            entity: {
                entityType: "dxc_customentityname",
                id: formContext.data.entity.getId(),
            },
            getMetadata: () => {
                return {
                    boundParameter: "entity",
                    operationName: actionName,
                    operationType: 0 /* Action */,
                    parameterTypes: {
                        entity: {
                            structuralProperty: 5, // Entity Type
                            typeName: "mscrm.dxc_customentityname",
                        },
                    },
                };
            },
        };

        try {
            Xrm.Utility.showProgressIndicator("Please wait...");
            const ret = await Xrm.WebApi.online.execute(dxcApproveRecord);
            if (ret.ok) {
                formContext.data.refresh(false);
            }
        } catch (error: any) {
            await Xrm.Navigation.openErrorDialog({
                details: `Error firing ${actionName} action: ${error.message}`,
                message: error.message,
            });
        } finally {
            Xrm.Utility.closeProgressIndicator();
        }
    }

    private static async setFormNotifications(formContext: Xrm.FormContext) {
        // NOTE: This shows a form notification if account status is disabled
        // This depends on this.statusReason being set in the onLoad event

        if (formContext.ui.getFormType() === XrmEnum.FormType.Update) {
            try {
                // Check status
                if (this.statusReason === AccountStatusCode.Inactive) {
                    formContext.ui.setFormNotification("Account is disabled!", "WARNING", "DISABLED");
                } else {
                    formContext.ui.clearFormNotification("DISABLED");
                }
            } catch (error: any) {
                await Xrm.Navigation.openErrorDialog({ message: `Error: ${error.message}` });
            }
        }
    }

    private static async setDecisionFields(formContext: Xrm.FormContext) {
        // NOTE: This function sets a user lookup and datetime field to capture who approved and when, using current datetime and logged in user
        setDecisionByOn(formContext, "dxc_approvedby", "dxc_approvedon");
    }
}

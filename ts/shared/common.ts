import { addYears } from "./datecalc";
import { SecurityRoles, userHasRoles } from "./security";

export function ViewFormSection(formContext: Xrm.FormContext, sectionname: string, visibility: boolean) {
    const tabs = formContext.ui.tabs.get();

    tabs.forEach((tab) => {
        tab.sections.forEach((section) => {
            if (section.getName() === sectionname) {
                section.setVisible(visibility);
            }
        });
    });
}

// Lock the passed list of fields on a form or editable grid
export function lockFields(formContext: Xrm.FormContext, disabledFields: string[]) {
    const currEntity = formContext.data.entity;
    currEntity.attributes.forEach((attribute) => {
        // Check if this attribute is in the passed in list of fields to be disabled
        if (disabledFields.indexOf(attribute.getName()) > -1) {
            // Disable all controls for this attribute, normally only one.
            attribute.controls.forEach((ctrl) => {
                (ctrl as Xrm.Controls.StringControl).setDisabled(true);
            });
        }
    });
}

// Lock the all fields on a from or editable grid except the passed list of fields
export function lockFieldsExcept(formContext: Xrm.FormContext, enabledFields: string[]) {
    const currEntity = formContext.data.entity;
    currEntity.attributes.forEach((attribute) => {
        // Check if this attribute is in the passed in list of enabled fields
        const enable = enabledFields.indexOf(attribute.getName()) >= 0;

        // Enable/Disable all controls for this attribute, normally only one.
        attribute.controls.forEach((ctrl) => {
            (ctrl as Xrm.Controls.StringControl).setDisabled(!enable);
        });
    });
}

/**
 * Detect when an editable grid has unsaved data
 * @param gridControl The grid control, normally returned by formContext.getControl("mygrididonform");
 */
export function gridIsDirty(gridControl: Xrm.Controls.GridControl): boolean {
    const rows = gridControl.getGrid().getSelectedRows();

    let isDirty = false;
    rows.forEach((row) => {
        const rowdata = row.data;

        // WARNING: rowdata.getIsDirty() is an undocumented function
        // it is also possible to use
        //     const qtyAttrib = rowdata.getAttribute("ebecs_requiredquantity")
        //     qtyAttrib.getIsDirty()
        // but this is unsupported too.
        if (typeof rowdata.getIsDirty !== "undefined" && rowdata.getIsDirty()) {
            isDirty = true;
        }
    });
    return isDirty;
}

/**
 * Return true if any field is dirty
 * @param logicalnames Array of logicalnames of fields that could be dirty
 */
export function anyFieldDirty(formContext: Xrm.FormContext, logicalnames: string[]): boolean {
    return logicalnames.some((field) => formContext.getAttribute(field)?.getIsDirty());
}

/**
 * Returns a promise that resolves in a the passed number of seconds
 */
export async function delay(ms: number) {
    return new Promise((r) => setTimeout(r, ms));
}

export function checkDateIsWithinPeriod(dateField: Xrm.Controls.DateControl, startDate: Date, endDate: Date) {
    const dateToCheck = dateField.getAttribute().getValue();

    // Build a dynamic ID incase this function is used on more than one control in the same form.
    // Use attribute name, because control name is not guaranteed to stay unchanged
    const ID_DATE = `idDate:${dateField.getAttribute().getName()}`;

    if (!dateToCheck) {
        dateField.clearNotification(ID_DATE);
        return "";
    }

    if (dateToCheck > endDate) {
        dateField.addNotification({
            messages: [`${dateField.getLabel()} must be on or before ${endDate.toLocaleDateString("en-GB")}`],
            notificationLevel: "ERROR",
            uniqueId: ID_DATE,
        });
        return ID_DATE;
    }

    if (dateToCheck < startDate) {
        dateField.addNotification({
            messages: [`${dateField.getLabel()} must be after ${startDate.toLocaleDateString("en-GB")}`],
            notificationLevel: "ERROR",
            uniqueId: ID_DATE,
        });
        return ID_DATE;
    }

    dateField.clearNotification(ID_DATE);
    return "";
}

export function checkDateWithinLastYear(dateField: Xrm.Controls.DateControl) {
    const now: Date = new Date();
    checkDateIsWithinPeriod(dateField, addYears(now, -1), now);
}

/*
 * Returns a lookup value with single element of the current user
 */
export function getCurrentUserLookup(): Xrm.LookupValue[] {
    const userSettings = Xrm.Utility.getGlobalContext().userSettings;
    return [
        {
            entityType: "systemuser",
            id: userSettings.userId,
            name: userSettings.userName,
        },
    ];
}

/*
 * Sets a pair of user and date/time attributes to the current user and current datestamp
 * @param formContext Xrm formcontext
 * @param userfield logical name of the user attribute to set to the current user
 * @param datefield logical name of the date attribute to set to the current date and time.
 */
export function setDecisionByOn(formContext: Xrm.FormContext, userfield: string, datefield: string) {
    const userAttrib: Xrm.Attributes.LookupAttribute = formContext.getAttribute(userfield);
    const dateAttrib: Xrm.Attributes.DateAttribute = formContext.getAttribute(datefield);

    userAttrib.setValue(getCurrentUserLookup());
    dateAttrib.setValue(new Date());
}

/*
 * Clears a pair of user and date/time attributes to the current user and current datestamp
 * @param formContext Xrm formcontext
 * @param userfield logical name of the user attribute to set to the current user
 */
export function clearDecisionByOn(formContext: Xrm.FormContext, userfield: string, datefield: string) {
    const userAttrib: Xrm.Attributes.LookupAttribute = formContext.getAttribute(userfield);
    const dateAttrib: Xrm.Attributes.DateAttribute = formContext.getAttribute(datefield);

    userAttrib.setValue(null);
    dateAttrib.setValue(null);
}

interface IError {
    message: string;
}

/*
 * Test if an variable of type 'unknown' contains a 'message' property.
 * All try/catch blocks provide an exception object that could be 'any' type, but it's now recommended
 * to use 'unknown' within catch() clauses, see
 * https://devblogs.microsoft.com/typescript/announcing-typescript-4-0/#unknown-on-catch
 */
export function hasMessage(obj: unknown): obj is IError {
    return obj !== null && typeof obj === "object" && typeof (obj as IError).message === "string";
}

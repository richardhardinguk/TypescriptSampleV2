import { createLookupValue } from "./webapi";

/**
 * Refresh the passed array of fields on the given form (formContext).
 * @param totalFields Array of fields to refresh.  Lookup fields must be provided in webapi style, '_ebecs_fieldid_value'.
 */
export async function refreshTotalFields(formContext: Xrm.FormContext, totalFields: string[]) {
    const TOTAL_FIELDS = totalFields;
    const entityName = formContext.data.entity.getEntityName();
    try {
        const lookupRegex = /^_(\w+)_value$/;
        const ret = await Xrm.WebApi.retrieveRecord(entityName, formContext.data.entity.getId(), "?$select=" + TOTAL_FIELDS.join(","));
        if (ret) {
            for (let field of TOTAL_FIELDS) {
                // See if the current field is a regex - if so need to rename to match the attrib,
                //  i.e   _ebecs_basketrateid_value becomes   ebecs_basketrateid
                const matchedLookup = lookupRegex.exec(field);

                const logicalName = matchedLookup ? matchedLookup[1] : field;
                const attrib: Xrm.Attributes.Attribute = formContext.getAttribute(logicalName);
                // Not all fields are on all forms, so test if attribute is found.
                if (attrib) {
                    if (matchedLookup) {
                        // Field is a lookup, so use createLookupValue build the lookup value
                        attrib.setValue(createLookupValue(ret, logicalName));
                    } else {
                        attrib.setValue(ret[logicalName]);
                    }
                }
            }
        }
    } catch (error: any) {
        await Xrm.Navigation.openErrorDialog({
            details: `Error refreshing total values on the ${entityName} form: ${error.message}`,
            message: error.message,
        });
    }
}

/**
 * Set the passed array of fields to never be submitted.  Normally same set of fields is also passed to refreshTotalFields()
 * @param totalFields Array of form fields. Lookup fields can be provided in webapi style, e.g. '_ebecs_fieldid_value' or 'ebecs_fieldid'
 */

export function neverSubmitTotalsFields(formContext: Xrm.FormContext, totalFields: string[]) {
    const TOTAL_FIELDS = totalFields;
    const lookupRegex = /^_(\w+)_value$/;

    for (let field of TOTAL_FIELDS) {
        const matchedLookup = lookupRegex.exec(field);

        const logicalName = (matchedLookup) ? matchedLookup[1] : field;
        const attrib: Xrm.Attributes.Attribute = formContext.getAttribute(logicalName);
        // Not all fields are on all forms, so test if attribute is found.
        if (attrib) {
            attrib.setSubmitMode("never");
        }
    }
}

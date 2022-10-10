/**
 * Take the output of a Xrm.WebApi.retrieve* and return an Xrm.LookupValue object
 * for easily setting a lookup value.
 * @param webApiResult the output of the WebApi call
 * @param field lookupfield included in the Xrm.WebApi.retrieve
 */
export function createLookupValue(webApiResult: { [key: string]: string }, field: string): Xrm.LookupValue[] | null {
    const guid = webApiResult[`_${field}_value`];
    if (guid === null) {
        return null;
    }
    return [
        {
            entityType: webApiResult[`_${field}_value@Microsoft.Dynamics.CRM.lookuplogicalname`],
            id: guid,
            name: webApiResult[`_${field}_value@OData.Community.Display.V1.FormattedValue`],
        },
    ];
}

export async function getEnvironmentVariableValue(schemaName: string, myXrm?: Xrm.XrmStatic): Promise<string> {
    if (typeof myXrm === "undefined") {
        myXrm = Xrm;
    }

    try {
        // prettier-ignore
        const fetchXml = [
            "<fetch>",
                "<entity name='environmentvariabledefinition'>",
                    "<attribute name='displayname'/>",
                    "<attribute name='defaultvalue'/>",
                    "<attribute name='schemaname'/>",
                    "<filter>",
                        `<condition attribute='schemaname' operator='eq' value='${schemaName}'/>`,
                    "</filter>",
                    "<link-entity name='environmentvariablevalue' from='environmentvariabledefinitionid'",
                                                                " to='environmentvariabledefinitionid' link-type='outer' alias='v'>",
                        "<attribute name='value'/>",
                    "</link-entity>",
                "</entity>",
            "</fetch>",
        ].join("");
        const ret = await myXrm.WebApi.online.retrieveMultipleRecords("environmentvariabledefinition", `?fetchXml=${fetchXml}`);

        if (ret && ret.entities && ret.entities.length === 1) {
            const row = ret.entities[0];

            // return the value, or default value or nothing
            return row["v.value"] || row.defaultvalue || "";
        }
        await myXrm.Navigation.openErrorDialog({ message: `Environment variable ${schemaName} does not exist` });
    } catch (error: any) {
        await myXrm.Navigation.openErrorDialog({ message: `Error getting environment Variable ${schemaName}: ${error.message}` });
    }
    return "";
}

/**
 * Lookup optionset values for a field. The results are cached using sessionStorage to reduce power platform requests.
 * @param entityLogicalname the logical name of the entity containing the attribute
 * @param attributeLogicalName the logical name of the attribute to be retreived
 */
export async function getMetadataOptionsetValues(entityLogicalname: string, attributeLogicalName: string): Promise<number[]> {
    // To reduce lookups to unchanging metadata we cache the results in SessionStorage, using the techniques described here.
    // https://community.dynamics.com/crm/b/crminthefield/posts/get-a-value-from-dynamics-365-ce-api-with-async-await-484252633?utm_source=dlvr.it&utm_medium=twitter
    // and
    // https://github.com/Ben-Fishin/Dynamics-365-CE-Client-Side-Scripting-Patterns/blob/14fbddcaa15a29addb843f18558a81c08d9592ff/solution/WebResources/dse_/Scripts/Common/common.js#L130

    const sessionStorageId = `${entityLogicalname}:${attributeLogicalName}`;

    // Check if browser has sessionStorage (97% of all browsers do as of 09/02/2022 according to caniuse.com)
    if (typeof Storage !== "undefined") {
        // Check if this value already exists in sessionStorage
        if (sessionStorage.getItem(sessionStorageId)) {
            // Return the result - and save a WebApi lookup
            return JSON.parse(sessionStorage[sessionStorageId]);
        }
    }

    try {
        const metadata = await Xrm.Utility.getEntityMetadata(entityLogicalname, [attributeLogicalName]);

        if (metadata.Attributes.getLength() !== 1) {
            await Xrm.Navigation.openErrorDialog({
                message: `No attribute found in metadata for table ${entityLogicalname} column ${attributeLogicalName}`,
            });
            return [];
        }

        const options: Xrm.Metadata.OptionMetadata[] = metadata.Attributes.get(0).OptionSet;

        // FIXME: As at 08/02/2022 the @types/xrm typings are incorrect here ('Value' instead of 'value'), so have to cast as any until updated.
        const optionsetValues = Object.values(options).map((x) => (x as any).value);

        // If the browser has support for Storage then save it in sessionStorage to save the lookup when queried again
        if (typeof Storage !== "undefined") {
            sessionStorage[sessionStorageId] = JSON.stringify(optionsetValues);
        }
        return optionsetValues;
    } catch {
        await Xrm.Navigation.openErrorDialog({ message: `Unable to get metadata for table ${entityLogicalname} column ${attributeLogicalName}` });
        return [];
    }
}

export const enum SecurityRoles {
    BasicUser = "Basic User",
    SystemAdministrator = "SystemAdministrator",
    CommercialAdministrator = "CommercialAdministrator" // custom role example
}

/**
 * Return true if the logged in users roles match with one of the passed role names.
 * @param rolesToCheck that contains the user roles name compared again the users' security roles
 * @param myXrm optional Xrm variable if only parent.Xrm is defined, such as used within an HTML webresource
 */
export function userHasRoles(rolesToCheck: string[], myXrm?: Xrm.XrmStatic): boolean {
    if (rolesToCheck.length < 1) {
        return false;
    }

    if (typeof myXrm === "undefined") {
        myXrm = Xrm;
    }

    // Get an array of GUIDs for the user roles that the current user if not passed in
    const userRoles = myXrm.Utility.getGlobalContext().userSettings.roles.get();

    // System Administrator gets all roles always
    if (userRoles.some((r) => r.name === "System Administrator")) {
        return true;
    }

    // Return true if any of the rolesToCheck are within the users current roles
    for (const checkName of rolesToCheck) {
        if (userRoles.some((r) => r.name === checkName)) {
            return true;
        }
    }
    return false;
}

/**
 * Return true if the logged in users is a member of the passed team.
 * @param teamName the name of the Team to check if the user is a member
 */
export async function userInTeams(teamNames: string[]): Promise<boolean> {
    if (teamNames.length < 1) {
        return false;
    }

    // Get the GUID of the user
    const userId = Xrm.Utility.getGlobalContext().userSettings.userId;

    // prettier-ignore
    const fetchXml = [
        "<fetch>",
            "<entity name='team' >",
                "<attribute name='name' />",
                "<link-entity name='teammembership' from='teamid' to='teamid' intersect='true' >",
                    "<filter>",
                        `<condition attribute='systemuserid' operator='eq' value='${userId}' />`,
                    "</filter>",
                "</link-entity>",
            "</entity>",
        "</fetch>"
    ].join("");

    try {
        const teams = await Xrm.WebApi.retrieveMultipleRecords("team", `?fetchXml=${fetchXml}`);

        // Return true if any of the teamNames are within the users teams returned by the WebAPI query
        for (const teamName of teamNames) {
            if (teams.entities.find((team) => team.name === teamName)) {
                return true;
            }
        }
    } catch (error: any) {
        await Xrm.Navigation.openErrorDialog({ message: `Error detecting team membership: ${error.message}` });
    }
    return false;
}

/**
 * Return true if the current user is a member of the passed team
 * @param teamId Guid of the team
 */
export async function userInTeam(teamId: string): Promise<boolean> {
    if (!teamId) {
        return false;
    }

    // Get the GUID of the user
    const userId = Xrm.Utility.getGlobalContext().userSettings.userId;

    try {
        const ret = await Xrm.WebApi.retrieveMultipleRecords("teammembership", `?$filter=teamid eq ${teamId} and systemuserid eq ${userId}`);
        if (ret) {
            if (ret.entities.length === 1) {
                return true;
            }
        }
    } catch (error: any) {
        await Xrm.Navigation.openErrorDialog({ message: `Error checking userInTeam: ${error.message}` });
    }
    return false;
}

/**
 * Set submit mode of passed array of fields to "Never" if the user doesn't have create & update of the field due to
 * FLS rules. This prevents the form from unnecessarily being set to dirty.
 * @param FLSFields Array of form fields
 */
export function setSubmitModeOnFLSFields(formContext: Xrm.FormContext, FLSfields: string[]) {
    for (const field of FLSfields) {
        const attrib: Xrm.Attributes.NumberAttribute = formContext.getAttribute(field);

        // Ensure field is on form
        if (attrib) {
            const canReadAndUpdate = attrib.getUserPrivilege().canRead && attrib.getUserPrivilege().canUpdate;

            if (!canReadAndUpdate) {
                attrib.setSubmitMode("never");
            }
        }
    }
}

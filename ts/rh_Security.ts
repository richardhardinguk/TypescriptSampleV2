import { SecurityRoles, userHasRoles, userInTeams, userInTeam } from "./shared/security"; // This imports a function from the shared security file

export class Security {
    public static userCommercialAdmin = false;

    public static onLoad(executionContext: Xrm.Events.EventContext) {
        // Check if user has role
        if (userHasRoles(["System Administrator"])) {
            // Do something
        }

        // Set in global variable onload then use elsewhere
        this.userCommercialAdmin = userHasRoles([SecurityRoles.CommercialAdministrator]);
    }

    // NOTE: Checks if user is member of the team that owns this record and return a promise
    // The userInTeams function requires passing the team name as a string
    public static async checkUserMemberOfOwnerTeam(formContext: Xrm.FormContext): Promise<boolean> {
        const ownerAttrib: Xrm.Attributes.Attribute = formContext.data.entity.attributes.get("ownerid");

        if (!ownerAttrib.getValue()[0]) {
            await Xrm.Navigation.openAlertDialog({
                text: `Not able to retrieve owner.`,
            });
        } else {
            const owner = ownerAttrib.getValue()[0];

            if (owner.entityType === "team") {
                return userInTeams([owner.name]);
            }
        }
        return false;
    }

    // NOTE: Checks if user is member of the team that owns this record and return a promise
    // The userInTeam function requires passing the team guid as a string
    public static async checkUserMemberOfTeamByGuid(executionContext: Xrm.Events.EventContext): Promise<boolean> {
        const formContext = executionContext.getFormContext();
        const teamAttrib: Xrm.Attributes.LookupAttribute = formContext.getAttribute("dxc_approvalteamid");

        if (!teamAttrib) {
            return false;
        }

        const approvalTeam = teamAttrib.getValue();
        if (!approvalTeam) {
            return false;
        }

        return userInTeam(approvalTeam[0].id);
    }
}

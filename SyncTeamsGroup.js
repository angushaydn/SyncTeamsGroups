var TeamSyncRibbon = TeamSyncRibbon || {};

TeamSyncRibbon.syncGroupMembersToCurrentTeam = function (primaryControl) {
	if (!primaryControl || !primaryControl.data || !primaryControl.data.entity) {
		return;
	}

	var entityName = primaryControl.data.entity.getEntityName();
	if (entityName !== "team") {
		return;
	}

	var teamType = primaryControl.getAttribute("teamtype").getValue();
	// Only sync if this is an O365 group enabled team
	if (teamType !== 3) {
		Xrm.Navigation.openAlertDialog({ text: "Only Office 365 Group enabled teams can sync group members." });
		return;
	}

	var rawId = primaryControl.data.entity.getId();
	if (!rawId) {
		Xrm.Navigation.openAlertDialog({ text: "Unable to find Team ID on this form." });
		return;
	}

	var teamId = rawId.replace(/[{}]/g, "");
	var executeRequest = {
		entity: { entityType: "team", id: teamId },
		getMetadata: function () {
			return {
				boundParameter: "entity",
				parameterTypes: {
					entity: { typeName: "mscrm.team", structuralProperty: 5 }
				},
				operationType: 0,
				operationName: "SyncGroupMembersToTeam"
			};
		}
	};

	Xrm.Utility.showProgressIndicator("Syncing group members to this team...");

	return Xrm.WebApi.execute(executeRequest)
		.then(function (response) {
			Xrm.Utility.closeProgressIndicator();
			if (response.ok) {
				return Xrm.Navigation.openAlertDialog({ text: "Team membership sync is successful." }).then(function () {
						var subgridControl = primaryControl.getControl("Members");
						if (subgridControl) {
							subgridControl.refresh();
						}
					}
				);
			}

			return Xrm.Navigation.openAlertDialog({ text: "Sync request did not complete successfully." });
		})
		.catch(function (error) {
			Xrm.Utility.closeProgressIndicator();
			return Xrm.Navigation.openErrorDialog({ message: error.message });
		});
};

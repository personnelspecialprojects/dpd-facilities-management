(window.FM_FILE_VERSIONS = window.FM_FILE_VERSIONS || {})["config.js"] = "2026.10.23-1";

/*
 * DPD Facilities Management: environment values, list names and column specs.
 *
 * This repository is public. Nothing that describes the building, the zones,
 * or when checks happen belongs in any file here. All of that is entered in
 * the app (Admin) and stored only in SharePoint.
 */
window.FM_CONFIG = {
  appName: "DPD Facilities Management",

  // Same Azure AD app registration as the Personnel Support portal
  // (PatrolBidPortal). These IDs aren't secrets.
  clientId: "0d384d2a-9a9e-46f6-a195-ad63f9e94f06",
  tenantId: "2935709e-c10c-4809-a302-852d369f8700",
  // Must match the SPA redirect URI registered in Azure exactly (trailing slash).
  redirectUri: "https://personnelspecialprojects.github.io/dpd-facilities-management/",

  // The SharePoint site itself: no trailing slash, no /SitePages/... part.
  siteUrl: "https://dallastxgov.sharepoint.com/sites/PatrolBid",
  loginScope: "User.Read",

  pollSeconds: 60,
  staleCheckSeconds: 180,

  // Watch labels only. No hours are kept anywhere in the app.
  watches: ["1st", "2nd", "3rd"],

  lists: {
    members: "FM Team Members",
    settings: "FM Settings",
    zones: "FM Zones",
    checks: "FM Check Log",
    taskTypes: "FM Task Types",
    activity: "FM Activity Log",
    work: "FM Work Items",
    workLog: "FM Work Item Log",
    audit: "FM Audit Log"
  },

  // Column specs used by Admin -> Check SharePoint setup.
  // type: Text (single line), Note (multiple lines, plain text), Number,
  //       Boolean (Yes/No), DateTime, Choice.
  // restricted: "required" = permissions must be limited to the team;
  //             "recommended" = should be.
  schema: {
    members: {
      restricted: "recommended",
      columns: [
        { name: "Title", type: "Text", note: "the person's work email" },
        { name: "DisplayName", type: "Text" },
        { name: "Role", type: "Choice", choices: ["Staff", "Admin"] },
        { name: "Active", type: "Boolean" }
      ]
    },
    settings: {
      restricted: "recommended",
      columns: [
        { name: "Title", type: "Text", note: "setting name" },
        { name: "Value", type: "Note" }
      ]
    },
    zones: {
      restricted: "required",
      columns: [
        { name: "Title", type: "Text", note: "zone name" },
        { name: "Description", type: "Note" },
        { name: "Watches", type: "Text" },
        { name: "SortOrder", type: "Number" },
        { name: "Active", type: "Boolean" }
      ]
    },
    checks: {
      restricted: "required",
      columns: [
        { name: "ZoneId", type: "Number" },
        { name: "ZoneName", type: "Text" },
        { name: "Watch", type: "Choice", choices: ["1st", "2nd", "3rd"] },
        { name: "BoardDay", type: "Text", indexed: true },
        { name: "CheckedAt", type: "DateTime", indexed: true },
        { name: "StaffEmail", type: "Text" },
        { name: "StaffName", type: "Text" },
        { name: "EnteredByEmail", type: "Text" },
        { name: "EnteredByName", type: "Text" },
        { name: "Method", type: "Choice", choices: ["Auto", "Manual", "Missed"] },
        { name: "Reason", type: "Note" },
        { name: "Voided", type: "Boolean" },
        { name: "VoidReason", type: "Note" },
        { name: "VoidedByEmail", type: "Text" }
      ]
    },
    taskTypes: {
      restricted: "recommended",
      columns: [
        { name: "Title", type: "Text", note: "task name" },
        { name: "InputType", type: "Choice", choices: ["Click Only", "Reference", "Count"] },
        { name: "InputLabel", type: "Text" },
        { name: "ChecklistSteps", type: "Note" },
        { name: "SortOrder", type: "Number" },
        { name: "Active", type: "Boolean" }
      ]
    },
    activity: {
      restricted: "recommended",
      columns: [
        { name: "TaskTypeId", type: "Number" },
        { name: "TaskTypeName", type: "Text" },
        { name: "Reference", type: "Text" },
        { name: "Quantity", type: "Number" },
        { name: "StaffEmail", type: "Text" },
        { name: "StaffName", type: "Text" },
        { name: "LoggedAt", type: "DateTime", indexed: true },
        { name: "Voided", type: "Boolean" },
        { name: "VoidReason", type: "Note" }
      ]
    },
    work: {
      restricted: "recommended",
      columns: [
        { name: "TaskTypeId", type: "Number" },
        { name: "TaskTypeName", type: "Text" },
        { name: "Description", type: "Note" },
        { name: "Reference", type: "Text" },
        { name: "Status", type: "Choice", choices: ["Open", "Completed", "Cancelled"], indexed: true },
        { name: "AssignedToEmail", type: "Text" },
        { name: "AssignedToName", type: "Text" },
        { name: "StepsState", type: "Note" },
        { name: "DueDate", type: "DateTime" },
        { name: "OpenedAt", type: "DateTime" },
        { name: "OpenedByName", type: "Text" },
        { name: "ClosedAt", type: "DateTime" }
      ]
    },
    workLog: {
      restricted: "recommended",
      columns: [
        { name: "WorkItemId", type: "Number", indexed: true },
        { name: "Details", type: "Note" },
        { name: "StaffEmail", type: "Text" },
        { name: "StaffName", type: "Text" }
      ]
    },
    audit: {
      restricted: "recommended",
      columns: [
        { name: "Area", type: "Text" },
        { name: "RecordId", type: "Text" },
        { name: "Details", type: "Note" },
        { name: "StaffEmail", type: "Text" }
      ]
    }
  }
};

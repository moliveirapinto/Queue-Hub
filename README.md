# Queue Hub - PCF Control for Dynamics 365 Customer Service

A **Power Apps Component Framework (PCF)** control designed for the **Dynamics 365 Customer Service workspace productivity pane**. It shows all queues the logged-in agent is assigned to and lets them drill into any queue to see members with their **real-time presence statuses**.

![Queue Hub](img/queue-hub.png)

## What It Does

This control provides a two-view navigation experience inside the productivity pane:

### Queue List View
- Displays all **queues the current agent is a member of** (via the `queuemembership` N:N relationship)
- **Search bar** to filter queues by name
- Initials-based icons for each queue
- Automatically **excludes system/internal queues** (e.g. auto-generated queue names)

### Agent Detail View
- Click any queue to see all **agents assigned to that queue**
- Each agent card shows:
  - **Avatar** with initials and a **presence dot** (color-coded)
  - **Real-time presence status** (Available, Busy, Away, Do Not Disturb, Offline, etc.)
  - **Duration** since last presence change
  - **"You" badge** highlighting the current agent
- **Summary chips** at the top showing a breakdown of presence statuses across the queue
- **Status-based sorting** — Available agents first, Offline agents last
- **10-second auto-refresh** — presence data polls every 10 seconds for near real-time updates

## Presence Status Colors

| Status | Color |
|---|---|
| Available | Green |
| Busy | Red |
| Do Not Disturb | Red |
| Away / Appear Away | Yellow |
| After Conversation Work | Pink |
| Offline / Inactive | Gray |

## Control Properties

| Property | Description | Default |
|---|---|---|
| **dummyProp** (SingleLine.Text) | Unused property required by PCF framework | — |

The control uses the **WebAPI** and **Utility** PCF features to query Dataverse directly.

## Prerequisites

- Dynamics 365 Customer Service with **Omnichannel for Customer Service** enabled
- **Productivity pane** configured in your Customer Service workspace app
- Agents must be **members of at least one queue** to see results
- Agent presence requires **Omnichannel presence** to be active

## How to Deploy to Your Dynamics 365 Environment

### Option 1: Import the Solution (Recommended)

1. Download the latest solution zip from the [Releases](../../releases) page
2. Go to your Dynamics 365 environment → **Settings** → **Solutions** (or use [make.powerapps.com](https://make.powerapps.com))
3. Click **Import** and upload the solution zip
4. Follow the import wizard and publish all customizations

### Option 2: Import via Power Platform CLI

```bash
# Install Power Platform CLI if not already installed
npm install -g pac

# Authenticate to your environment
pac auth create --url https://YOUR_ORG.crm.dynamics.com

# Import the solution
pac solution import --path ./solution.zip
```

### Option 3: Build from Source

If you want to modify the control and rebuild:

```bash
# Clone the repository
git clone https://github.com/moliveirapinto/Queue-Hub.git
cd Queue-Hub/QueueHub

# Install dependencies
npm install

# Build the control
npm run build

# Create a solution project
mkdir ../Solution && cd ../Solution
pac solution init --publisher-name yourpublisher --publisher-prefix yourprefix
pac solution add-reference --path ../QueueHub

# Build the solution zip (unmanaged)
dotnet build --configuration Debug

# Import to your environment
pac solution import --path bin/Debug/Solution.zip
```

## Configure as a Productivity Pane Tool

After importing the solution, you need to configure Queue Hub as a **pane tool** in the productivity pane:

1. Open [make.powerapps.com](https://make.powerapps.com) and navigate to your environment
2. Go to **Apps** → open **Customer Service admin center**
3. Navigate to **Workspaces** → **Productivity pane**
4. Under **Pane tools**, click **+ Add tool**
5. Configure the pane tool:
   - **Name**: Queue Hub
   - **Unique name**: queue_hub
   - **Control name**: `maulabs_MauLabs.QueueHub`
   - **Global**: Yes (available on all sessions)
6. Save and enable the tool
7. Add it to your **Pane tab configuration** linked to your productivity pane config
8. **Publish** all customizations

The control will appear as a new icon in the productivity pane sidebar of the Customer Service workspace.

![Pane Tool Configuration](img/pane-tool.png)

## Project Structure

```
Queue-Hub/
├── README.md
├── SyncPhotos.ps1                             # PowerShell script for manual photo sync
├── SyncUserPhotos-flow-definition.json        # Power Automate flow definition (reference)
├── img/                                       # Screenshots for documentation
└── QueueHub/
    ├── package.json
    ├── tsconfig.json
    ├── pcfconfig.json
    ├── QueueHub.pcfproj
    └── QueueHub/
        ├── ControlManifest.Input.xml          # PCF control manifest
        ├── index.ts                           # Main control logic (~470 lines)
        └── css/
            └── QueueHub.css                   # Control styles
```

## Profile Photo Sync

The Queue Hub displays agent profile photos from the Dataverse `entityimage` field. Since Dataverse doesn't automatically sync photos from Azure AD / Microsoft Entra ID, you need to sync them periodically.

### Option 1: Power Automate Cloud Flow (Automated — Recommended)

A cloud flow named **"Sync User Photos"** is included in the QueueHub solution. It automatically syncs profile photos from Microsoft 365 (Office 365 Users connector) to the Dataverse `systemuser.entityimage` field.

#### How It Works

```
Recurrence (Daily 06:00 UTC)
  └─ List Active Users (Dataverse)
       Filter: enabled, has AAD Object ID, excludes application/non-interactive users
       └─ For Each User (sequential, concurrency=1)
            └─ Scope: Sync Photo
                 ├─ Get User Photo (Office 365 Users connector)
                 └─ Update User Record — sets entityimage (Dataverse)
```

**Flow actions in detail:**

| Step | Action | Connector | Description |
|------|--------|-----------|-------------|
| 1 | **Recurrence** | Built-in | Triggers daily at 06:00 UTC |
| 2 | **List_Active_Users** | Dataverse | Queries `systemusers` with filter: `isdisabled eq false and azureactivedirectoryobjectid ne null and accessmode ne 4 and accessmode ne 6` — this returns only real human agents, excluding disabled users, bot accounts, and Copilot/application service principals |
| 3 | **Apply_to_each_user** | Control | Iterates through each user sequentially (concurrency=1 to avoid throttling) |
| 4 | **Scope_Sync_Photo** | Control | Wraps the photo sync steps so that a failure for one user (e.g. no photo in M365) doesn't stop processing of other users |
| 5 | **Get_User_Photo** | Office 365 Users | Calls `UserPhoto` operation with the user's `azureactivedirectoryobjectid` to fetch the profile photo from Microsoft 365 |
| 6 | **Update_User_Photo** | Dataverse | Updates the `systemuser` record's `entityimage` field with the photo content from step 5 |

#### Connection References

The flow uses two connection references that must be configured in your environment:

| Connection Reference | Connector | Purpose |
|---|---|---|
| `msdyn_sharedcommondataserviceforapps_*` | Microsoft Dataverse | Read/write systemuser records |
| `maulabs_sharedoffice365users_*` | Office 365 Users | Read user profile photos from Microsoft 365 |

#### Setup Instructions

1. Go to [make.powerapps.com](https://make.powerapps.com) → **Solutions** → **QueueHub**
2. Open the flow **"Sync User Photos"**
3. Set up the required connections:
   - **Dataverse** — select your environment connection
   - **Office 365 Users** — sign in with an account that has permissions to read user photos
4. **Turn on** the flow
5. Optionally click **Run** to trigger an immediate sync

#### Important Notes

- The filter `accessmode ne 4 and accessmode ne 6` excludes **Non-Interactive** (4) and **Application** (6) system users — these are Copilot agents, bots, and service principals that don't have real profile photos and would cause `BadRequest` errors when attempting to update their records.
- The **Scope** pattern ensures that if fetching a photo fails for one user (e.g. user has no photo in Microsoft 365), the flow continues processing the remaining users.
- Concurrency is set to **1** (sequential processing) to avoid API throttling.
- The flow definition JSON is available in [`SyncUserPhotos-flow-definition.json`](SyncUserPhotos-flow-definition.json) for reference. You can use this as a template to recreate the flow in another environment.

### Option 2: PowerShell Script (Manual / Scheduled)

Run `SyncPhotos.ps1` from a PowerShell terminal for on-demand photo sync:

```powershell
.\SyncPhotos.ps1
```

The script uses device code flow for authentication (no app registration required) and syncs photos from Microsoft Graph to Dataverse `entityimage` for all active users with Azure AD Object IDs.

You can schedule it via **Windows Task Scheduler** or **Azure Automation** for periodic runs.

## Data Model

The control queries three main entities via FetchXML:

```
queue ←→ queuemembership (N:N) ←→ systemuser
                                        ↓
                                  msdyn_agentstatus
                                        ↓
                                  msdyn_presence
```

- **Queues**: Filtered to only show queues where the current user is a member
- **Agents**: Retrieved per queue via the `queuemembership` intersect entity
- **Presence**: Loaded from `msdyn_agentstatus` (batched in groups of 10) and resolved to friendly names via `msdyn_presence`

## License

MIT

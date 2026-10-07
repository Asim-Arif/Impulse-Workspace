# IntraOffice Collaboration & CRM Module - Developer Integration Guide

**Target System:** Impulse ERP Ecosystem (.NET 8 / Blazor Server / SQL Server)  
**Module:** IntraOffice (Real-time Messaging, Channels, Tasks, Approvals/Minutes, Meetings, CRM & Leads, Sticky Notes, Presence)  
**Audience:** Software Engineers & Solution Architects integrating IntraOffice into sibling ERP suites (e.g., Textile, Manufacturing, Surgical).

---

## 1. Architectural Overview

The IntraOffice module is structured as a decoupled, multi-tier subsystem that easily plugs into any ASP.NET Core Blazor application sharing the core ERP database foundation.

```
┌────────────────────────────────────────────────────────┐
│                   Blazor UI Pages                      │
│ (Chat, Channels, Tasks, Minutes, Meetings, Leads, CRM) │
└──────────────────────────┬─────────────────────────────┘
                           │
┌──────────────────────────▼─────────────────────────────┐
│                 Service Layer (Impulse)                │
│    IIntraOfficeService, Adapters, SignalR Hubs         │
└──────────────────────────┬─────────────────────────────┘
                           │
┌──────────────────────────▼─────────────────────────────┐
│               Data Access Layer (Dapper)               │
│          IIntraOfficeDataAccess & Models               │
└──────────────────────────┬─────────────────────────────┘
                           │
┌──────────────────────────▼─────────────────────────────┐
│              SQL Server (Idempotent DDL)               │
│      18+ IntraOffice Tables + Core System Joins        │
└────────────────────────────────────────────────────────┘
```

---

## 2. Package Folder Structure

This package has been organized to match the target solution layout:

```
Code Sharing Interaoffice/
├── Database/
│   └── 01_IntraOffice_Tables_And_Schema.sql     <-- Idempotent SQL Server table & seed script
├── DataAccessLibrary/
│   ├── DAC/
│   │   └── IntraOfficeDataAccess.cs             <-- Dapper-based high performance async DAC
│   ├── Interface/
│   │   └── IIntraOfficeDataAccess.cs            <-- Data Access contract
│   └── Models/                                  <-- Strongly-typed DTOs & Models
│       ├── EmailModels.cs
│       ├── IntraAnnouncementModels.cs
│       ├── IntraChannelModels.cs
│       ├── IntraCrmDashboardModels.cs
│       ├── IntraCustomer360Models.cs
│       ├── IntraEnums.cs
│       ├── IntraLeadModels.cs
│       ├── IntraMeetingModels.cs
│       ├── IntraMessageModels.cs
│       ├── IntraMinutesModels.cs
│       ├── IntraPresenceAndNotesModels.cs
│       ├── IntraTaskModels.cs
│       ├── IntraUserModels.cs
│       └── IntraWorkflowModels.cs
├── Impulse/
│   ├── Hubs/
│   │   └── ChatHub.cs                           <-- Real-time SignalR hub for chat & presence
│   ├── Pages/
│   │   └── IntraOffice/                         <-- Blazor UI components & dashboard views
│   │       ├── Admin/                           <-- Email configuration, minute types admin
│   │       ├── AiAssistant/                     <-- AI assistant integration page
│   │       ├── Announcements/                   <-- Company notice board
│   │       ├── Channels/                        <-- Slack/Teams-style discussion channels
│   │       ├── Chat/                            <-- Direct messaging & group chat rooms
│   │       ├── CRM/                             <-- Customer 360, Leads pipeline, Email templates
│   │       ├── Directory/                       <-- Interactive employee directory
│   │       ├── Meetings/                        <-- Scheduled & live meeting management
│   │       ├── Minutes/                         <-- Multi-step approval workflows
│   │       ├── StickyNotes/                     <-- Floating user sticky notes
│   │       ├── Tasks/                           <-- Kanban & list task management
│   │       ├── OfficeDashboard.razor            <-- Master IntraOffice portal
│   │       └── OfficeHub.razor
│   ├── Services/
│   │   └── IntraOffice/                         <-- Business logic, notifications, adapters
│   │       ├── IIntraOfficeService.cs
│   │       ├── IntraOfficeService.cs
│   │       ├── IntraOfficeAdapters.cs           <-- Clean adapter facade for modular UI consumption
│   │       ├── AppNotificationService.cs        <-- Realtime in-app sound/banner events
│   │       ├── EmailEncryptionService.cs        <-- AES credential security for SMTP
│   │       ├── FileService.cs                   <-- Upload streaming & MIME handling
│   │       └── ...
│   └── wwwroot/
│       ├── css/
│       │   └── intraoffice.css                  <-- Tailored styles, chat bubbles, Kanban, cards
│       └── js/
│           ├── intraoffice.js                   <-- DOM utilities, auto-scroll, audio player
│           ├── crm-charts.js                    <-- Chart.js bridges for CRM analytics
│           ├── audioRecorder.js                 <-- Voice note capture via Web Audio API
│           └── audio.js
└── DEVELOPER_INTEGRATION_GUIDE.md               <-- This documentation
```

---

## 3. Database Deployment

Run the included migration script on the target product database (e.g., `Impulse_Textile`):

- **Script Path:** `Database/01_IntraOffice_Tables_And_Schema.sql`
- **Safety:** 100% idempotent (`IF NOT EXISTS` guards everywhere). It will not drop or truncate existing data.
- **Created Tables:**
  - `UserPresences`
  - `Channels`, `ChannelMembers`
  - `Messages`, `MessageAttachments`
  - `Announcements`, `AnnouncementAcknowledgments`, `AnnouncementAttachments`
  - `TaskItems`, `Task_Assignees`, `Task_Roles`, `TaskComments`, `TaskAttachments`
  - `Meetings`, `MeetingParticipants`
  - `MinuteTypes`, `MinuteApprovals`, `MinuteWorkflowHistories`, `MinuteAttachments`
  - `StickyNotes`
  - `Leads`, `LeadActivities`
  - `EmailConfigurations`, `EmailTemplates`

### Core ERP Schema Assumptions
The DAC assumes standard core tables are present in the target database:
- `Users` (`UserID`, `UserName`, `FullUserName`, `EmpID`, `InActive`)
- `Employees` (`empid`, `name`, `Designation`, `deptid`, `Phone1`)
- `Departments` (`deptid`, `name`)
- `ForeignCustomers` & `FCustomerOrders` *(Optional for CRM / Customer 360 lookup; if not used in Textile, these queries can be adapted)*

---

## 4. ASP.NET Core `Program.cs` Setup

### 4.1 Dependency Injection Registrations
Add the following service registrations to your `Program.cs` (or `Startup.cs`):

```csharp
// =============================================================
// IntraOffice Module Registrations
// =============================================================
builder.Services.AddScoped<DataAccessLibrary.Interface.IntraOffice.IIntraOfficeDataAccess, 
                           DataAccessLibrary.DAC.IntraOffice.IntraOfficeDataAccess>();
builder.Services.AddScoped<Impulse.Services.IntraOffice.IIntraOfficeService, 
                           Impulse.Services.IntraOffice.IntraOfficeService>();

// Notifications & Adapters
builder.Services.AddSingleton<Impulse.Services.IntraOffice.IAppNotificationService, 
                              Impulse.Services.IntraOffice.AppNotificationService>();
builder.Services.AddSingleton<Impulse.Services.IntraOffice.StickyNoteNotificationService>();
builder.Services.AddSingleton<Impulse.Services.IntraOffice.IEmailEncryptionService, 
                              Impulse.Services.IntraOffice.EmailEncryptionService>();

// File Storage & Module Services
builder.Services.AddScoped<Impulse.Services.IntraOffice.IFileService, 
                           Impulse.Services.IntraOffice.FileService>();
builder.Services.AddScoped<Impulse.Services.IntraOffice.IChannelService, 
                           Impulse.Services.IntraOffice.ChannelService>();
builder.Services.AddScoped<Impulse.Services.IntraOffice.IMessageService, 
                           Impulse.Services.IntraOffice.MessageService>();
builder.Services.AddScoped<Impulse.Services.IntraOffice.IAnnouncementService, 
                           Impulse.Services.IntraOffice.AnnouncementService>();
builder.Services.AddScoped<Impulse.Services.IntraOffice.ITaskService, 
                           Impulse.Services.IntraOffice.TaskService>();
builder.Services.AddScoped<Impulse.Services.IntraOffice.IMeetingService, 
                           Impulse.Services.IntraOffice.MeetingService>();
builder.Services.AddScoped<Impulse.Services.IntraOffice.IMinuteApprovalService, 
                           Impulse.Services.IntraOffice.MinuteApprovalService>();
builder.Services.AddScoped<Impulse.Services.IntraOffice.IMinuteTypeService, 
                           Impulse.Services.IntraOffice.MinuteTypeService>();
builder.Services.AddScoped<Impulse.Services.IntraOffice.IEmailService, 
                           Impulse.Services.IntraOffice.EmailService>();
builder.Services.AddScoped<Impulse.Services.IntraOffice.ICrmService, 
                           Impulse.Services.IntraOffice.CrmService>();

// Optional Integrations (AI / WhatsApp)
builder.Services.AddScoped<Impulse.Services.IntraOffice.IAiAssistantService, 
                           Impulse.Services.IntraOffice.AiAssistantService>();
builder.Services.AddHttpClient<Impulse.Services.IntraOffice.IAiService, 
                              Impulse.Services.IntraOffice.OpenRouterAiService>();
builder.Services.AddHttpClient<Impulse.Services.IntraOffice.IWhatsAppNotificationService, 
                              Impulse.Services.IntraOffice.WhatsAppNotificationService>();
```

### 4.2 SignalR Real-Time Hub Mapping
In the endpoint configuration section of `Program.cs`:

```csharp
app.MapBlazorHub();
app.MapHub<Impulse.Hubs.ChatHub>("/chathub");
```

### 4.3 Upload Directories Initialization
Ensure upload subdirectories are created on application startup so attachment uploads do not throw `DirectoryNotFoundException`:

```csharp
// Ensure IntraOffice upload directories exist
var env = app.Services.GetRequiredService<IWebHostEnvironment>();
var webRoot = env.WebRootPath ?? Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
foreach (var sub in new[] { "tasks", "announcements", "chat", "meetings", "leads", "minutes" })
{
    var dir = Path.Combine(webRoot, "uploads", sub);
    if (!Directory.Exists(dir)) 
    {
        Directory.CreateDirectory(dir);
    }
}
```

---

## 5. UI & Static Asset Integration

### 5.1 Host Page Assets (`_Host.cshtml` or `App.razor`)
Ensure the following CSS and script tags are included:

```html
<!-- IntraOffice Stylesheet -->
<link href="css/intraoffice.css" rel="stylesheet" />

<!-- Chart.js (Required for CRM analytics & metrics) -->
<script src="js/chart.min.js"></script>

<!-- IntraOffice Scripts -->
<script src="js/intraoffice.js"></script>
<script src="js/crm-charts.js"></script>
<script src="js/audioRecorder.js"></script>
<script src="js/audio.js"></script>
```

### 5.2 Navigation Menu Integration
Add navigation items to your `NavMenu.razor` under a dedicated **IntraOffice** or **Collaboration** group:

```razor
<li class="nav-item">
    <NavLink class="nav-link" href="intraoffice/dashboard">
        <i class="fas fa-building me-2"></i> IntraOffice Hub
    </NavLink>
</li>
<li class="nav-item">
    <NavLink class="nav-link" href="intraoffice/chat">
        <i class="fas fa-comments me-2"></i> Team Chat
    </NavLink>
</li>
<li class="nav-item">
    <NavLink class="nav-link" href="intraoffice/tasks">
        <i class="fas fa-tasks me-2"></i> Tasks & Projects
    </NavLink>
</li>
<li class="nav-item">
    <NavLink class="nav-link" href="intraoffice/minutes">
        <i class="fas fa-file-signature me-2"></i> Approvals & Minutes
    </NavLink>
</li>
<li class="nav-item">
    <NavLink class="nav-link" href="intraoffice/leads">
        <i class="fas fa-funnel-dollar me-2"></i> Leads & CRM
    </NavLink>
</li>
```

---

## 6. Key Integration Tips for the Developer

1. **Authentication Context:**
   The components fetch the active user identity from `AuthenticationStateProvider` (falling back to username/session if standard ASP.NET Identity or custom cookie authentication is used).
2. **Dapper Queries:**
   All queries inside `IntraOfficeDataAccess.cs` utilize parameterized Dapper calls (`QueryAsync`, `ExecuteAsync`, `ExecuteScalarAsync`) to ensure maximum throughput and protection against SQL injection.
3. **Database Portability:**
   No table names are hardcoded with database prefixes. All queries run against the current connection configured under `ConnectionStrings:DefaultConnection`.
4. **CRM & Customer 360 Customization:**
   If the Textile product uses custom customer or inventory tables (e.g. `TextileCustomers` instead of `ForeignCustomers`), modify the read queries in `Customer360` methods in `IntraOfficeDataAccess.cs` to bind to your specific ERP tables.

---

*Package prepared for Impulse Textile Suite Integration. For questions or modifications, refer to `IIntraOfficeDataAccess.cs` and `IntraOfficeAdapters.cs`.*

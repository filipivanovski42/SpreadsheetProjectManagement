# Spreadsheet Project Management System

A Google Apps Script-based project management system designed for sprint-based workflow with parallel workstreams.

## Overview

This project provides a comprehensive spreadsheet-based project management solution that allows teams to organize work into sprints and parallel workstreams, with automated dashboards and task tracking capabilities.

## Features

### 📊 Sprint & Workstream Organization
- **Project Structure**: Each project is its own spreadsheet
- **Parallel Workstreams**: Each sheet represents a different sprint or workstream (e.g., marketing, core development)
- **Dashboard-First View**: Each sheet displays a dashboard showing overall completion status

### 📈 Dashboard Analytics
- **Time Tracking**: Total estimated time vs actual time spent
- **Progress Visualization**: Progress bars and sparklines showing overall project completion
- **Real-time Updates**: Automatic calculation of project metrics

### 🚀 Automated Sprint Management
- **Sprint Initialization**: Apps Script wizard to create new sprints/workstreams
- **Team Management**: Prompt-based team member addition with dropdown integration
- **Custom Naming**: Flexible naming for sprints and workstreams

### ✅ Task Management
- **Quick Task Creation**: One-click task creation with detail prompts
- **Permission Control**: Users can only complete tasks assigned to them
- **Email Notifications**: Optional email notifications when tasks are assigned

### ⏱️ Time & Deadline Tracking
- **Time Estimation**: Estimated hours vs actual hours tracking
- **Deadline Management**: Automatic highlighting of overdue tasks in red
- **Completion Requirements**: Users must input hours spent before marking tasks complete

### 🎨 Status Management & Visual Indicators
- **Status Options**: Not Started, In Progress, Blocked, Done
- **Color Coding**:
  - 🔴 **Blocked**: Red background
  - 🟢 **Done**: Green background  
  - 🟡 **In Progress**: Yellow background
  - ⚫ **Not Started**: Gray background
- **Pleasant Design**: Desaturated pastel color palette for easy viewing

## Files

- `ProductionScript.gs` - Main Google Apps Script implementation
- `ProductionNew.txt` - Project requirements and specifications

## Setup

1. Create a new Google Spreadsheet
2. Open the Apps Script editor
3. Copy the contents of `ProductionScript.gs`
4. Configure the script with your team members and project settings
5. Run the initialization functions to set up your first sprint

## Usage

### Creating a New Sprint
1. Run the "Initialize Sprint" function from the Apps Script editor
2. Enter the sprint/workstream name
3. Add team members who will be working on this sprint
4. The script will automatically create the sheet structure and dashboard

### Adding Tasks
1. Click the "Create New Task" button
2. Fill in task details including:
   - Task name and description
   - Assignee (from dropdown)
   - Estimated hours
   - Deadline
   - Initial status
3. Optionally send email notification to assignee

### Updating Task Progress
1. Team members can only update tasks assigned to them
2. Input actual hours spent
3. Update status as work progresses
4. Mark as complete when finished

## Permissions & Security

- Google account integration for user authentication
- Role-based permissions (task creators vs assignees)
- Secure email notification system
- Controlled access to task completion functions

## Contributing

This project is designed for internal team use. For feature requests or bug reports, please contact the project maintainer.

## License

Internal use only.

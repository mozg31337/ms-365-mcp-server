# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.10.0] - 2025-06-30

### Added

- **Microsoft Planner Plans Discovery**: Added support for listing and searching planner plans
  - `list-user-planner-plans`: List all planner plans for the current user with advanced filtering
  - Support for OData query parameters: `$filter`, `$search`, `$select`, `$orderby`, `$top`, `$skip`, `$count`
  - Plan filtering by title: `$filter=contains(title,'IT')`
  - Plan search functionality: `$search=IT department`
  - Enhanced workflow: "Get tasks from plan named 'IT tasks'" now supported
- **Microsoft Planner Task Management**: Added support for updating and deleting planner tasks
  - `update-planner-task`: Update properties of existing planner tasks with ETag support for concurrency control
  - `delete-planner-task`: Delete planner tasks with ETag support for concurrency control
- **Critical Bug Fix**: Fixed ETag handling for planner task updates
  - `@odata.etag` headers now preserved in GET responses
  - Resolves 400 "invalid format" and 412 "precondition failed" errors
- ETag support for planner task operations to prevent concurrent modification conflicts
- Comprehensive test coverage for new planner operations (37 total tests, 25 new)
- Support for all updatable planner task properties (title, percentComplete, priority, dueDateTime, assignments, appliedCategories, bucketId)

### Enhanced User Experience
- **Natural Language Queries**: Users can now say "Give me all tasks from the planner named 'IT tasks'"
- **Plan Discovery**: No longer need to know plan IDs - search by name or list all plans
- **Advanced Filtering**: Rich query capabilities for finding specific plans

### Technical Details

- Added proper `If-Match` header support for Microsoft Graph API concurrency control
- Error handling for 409 (conflict) and 412 (precondition failed) scenarios
- Operations automatically disabled in read-only mode
- Full backward compatibility with existing functionality

### Dependencies

- Microsoft Graph API permissions: `Tasks.Read` scope for plan listing and search, `Tasks.ReadWrite` scope required for update/delete operations

## [0.9.12] - Previous Release

- Previous functionality and features

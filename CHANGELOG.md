# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.10.0] - 2025-06-30

### Added

- **Microsoft Planner Task Management**: Added support for updating and deleting planner tasks
  - `update-planner-task`: Update properties of existing planner tasks with ETag support for concurrency control
  - `delete-planner-task`: Delete planner tasks with ETag support for concurrency control
- ETag support for planner task operations to prevent concurrent modification conflicts
- Comprehensive test coverage for new planner task operations
- Support for all updatable planner task properties (title, percentComplete, priority, dueDateTime, assignments, appliedCategories, bucketId)

### Technical Details

- Added proper `If-Match` header support for Microsoft Graph API concurrency control
- Error handling for 409 (conflict) and 412 (precondition failed) scenarios
- Operations automatically disabled in read-only mode
- Full backward compatibility with existing functionality

### Dependencies

- Microsoft Graph API permissions: `Tasks.ReadWrite` scope required for update/delete operations

## [0.9.12] - Previous Release

- Previous functionality and features

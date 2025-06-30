import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { registerGraphTools } from '../src/graph-tools.js';
import GraphClient from '../src/graph-client.js';

// Mock GraphClient
vi.mock('../src/graph-client.js');

describe('Planner Task Operations', () => {
  let server: McpServer;
  let mockGraphClient: vi.Mocked<GraphClient>;

  beforeEach(() => {
    vi.clearAllMocks();

    // Create mock server
    server = {
      tool: vi.fn(),
    } as any;

    // Create mock GraphClient instance
    mockGraphClient = {
      graphRequest: vi.fn(),
    } as any;

    // Mock the GraphClient constructor
    (GraphClient as any).mockImplementation(() => mockGraphClient);
  });

  afterEach(() => {
    vi.resetAllMocks();
  });

  describe('update-planner-task', () => {
    beforeEach(() => {
      // Register tools to set up the update-planner-task tool
      registerGraphTools(server, mockGraphClient, false);
    });

    it('should be registered as a tool', () => {
      expect(server.tool).toHaveBeenCalledWith(
        'update-planner-task',
        expect.any(String),
        expect.any(Object),
        expect.any(Object),
        expect.any(Function)
      );
    });

    it('should handle successful update with If-Match header', async () => {
      // Mock successful response
      mockGraphClient.graphRequest.mockResolvedValue({
        content: [{ text: '{}' }],
        isError: false,
      });

      // Find the update tool handler
      const updateToolCall = (server.tool as any).mock.calls.find(
        (call: any) => call[0] === 'update-planner-task'
      );
      expect(updateToolCall).toBeDefined();

      const [, , , , handler] = updateToolCall;

      // Test the handler with proper parameters
      const result = await handler({
        plannerTaskId: 'task-123',
        'If-Match': 'W/"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBAWCc="',
        body: JSON.stringify({
          title: 'Updated Task Title',
          percentComplete: 50,
        }),
      });

      expect(mockGraphClient.graphRequest).toHaveBeenCalledWith(
        '/planner/tasks/task-123',
        expect.objectContaining({
          method: 'PATCH',
          headers: expect.objectContaining({
            'If-Match': 'W/"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBAWCc="',
          }),
          body: expect.stringContaining('Updated Task Title'),
        })
      );

      expect(result.content).toHaveLength(1);
      expect(result.isError).toBeFalsy();
    });

    it('should handle 409 conflict error when ETag is outdated', async () => {
      // Mock conflict response
      mockGraphClient.graphRequest.mockResolvedValue({
        content: [{ text: JSON.stringify({ error: 'Conflict - resource has been modified' }) }],
        isError: true,
      });

      const updateToolCall = (server.tool as any).mock.calls.find(
        (call: any) => call[0] === 'update-planner-task'
      );
      const [, , , , handler] = updateToolCall;

      const result = await handler({
        plannerTaskId: 'task-123',
        'If-Match': 'W/"outdated-etag"',
        body: JSON.stringify({ title: 'Updated Title' }),
      });

      expect(result.isError).toBeTruthy();
      expect(result.content[0].text).toContain('Conflict');
    });

    it('should handle missing If-Match header', async () => {
      const updateToolCall = (server.tool as any).mock.calls.find(
        (call: any) => call[0] === 'update-planner-task'
      );
      const [, , , , handler] = updateToolCall;

      // Test without If-Match header - should still work but may fail at API level
      mockGraphClient.graphRequest.mockResolvedValue({
        content: [{ text: JSON.stringify({ error: 'If-Match header is required' }) }],
        isError: true,
      });

      const result = await handler({
        plannerTaskId: 'task-123',
        body: JSON.stringify({ title: 'Updated Title' }),
      });

      expect(result.isError).toBeTruthy();
    });
  });

  describe('delete-planner-task', () => {
    beforeEach(() => {
      // Register tools to set up the delete-planner-task tool
      registerGraphTools(server, mockGraphClient, false);
    });

    it('should be registered as a tool', () => {
      expect(server.tool).toHaveBeenCalledWith(
        'delete-planner-task',
        expect.any(String),
        expect.any(Object),
        expect.any(Object),
        expect.any(Function)
      );
    });

    it('should handle successful deletion with If-Match header', async () => {
      // Mock successful deletion (204 No Content)
      mockGraphClient.graphRequest.mockResolvedValue({
        content: [{ text: '' }],
        isError: false,
      });

      const deleteToolCall = (server.tool as any).mock.calls.find(
        (call: any) => call[0] === 'delete-planner-task'
      );
      expect(deleteToolCall).toBeDefined();

      const [, , , , handler] = deleteToolCall;

      const result = await handler({
        plannerTaskId: 'task-123',
        'If-Match': 'W/"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBAWCc="',
      });

      expect(mockGraphClient.graphRequest).toHaveBeenCalledWith(
        '/planner/tasks/task-123',
        expect.objectContaining({
          method: 'DELETE',
          headers: expect.objectContaining({
            'If-Match': 'W/"JzEtVGFzayAgQEBAQEBAQEBAQEBAQEBAWCc="',
          }),
        })
      );

      expect(result.isError).toBeFalsy();
    });

    it('should handle 404 not found error', async () => {
      mockGraphClient.graphRequest.mockResolvedValue({
        content: [{ text: JSON.stringify({ error: 'Task not found' }) }],
        isError: true,
      });

      const deleteToolCall = (server.tool as any).mock.calls.find(
        (call: any) => call[0] === 'delete-planner-task'
      );
      const [, , , , handler] = deleteToolCall;

      const result = await handler({
        plannerTaskId: 'nonexistent-task',
        'If-Match': 'W/"some-etag"',
      });

      expect(result.isError).toBeTruthy();
      expect(result.content[0].text).toContain('Task not found');
    });

    it('should handle 412 precondition failed error', async () => {
      mockGraphClient.graphRequest.mockResolvedValue({
        content: [
          {
            text: JSON.stringify({
              error: 'Precondition failed - If-Match header does not match current ETag',
            }),
          },
        ],
        isError: true,
      });

      const deleteToolCall = (server.tool as any).mock.calls.find(
        (call: any) => call[0] === 'delete-planner-task'
      );
      const [, , , , handler] = deleteToolCall;

      const result = await handler({
        plannerTaskId: 'task-123',
        'If-Match': 'W/"wrong-etag"',
      });

      expect(result.isError).toBeTruthy();
      expect(result.content[0].text).toContain('Precondition failed');
    });
  });

  describe('read-only mode', () => {
    it('should skip update and delete tools in read-only mode', () => {
      const readOnlyServer = {
        tool: vi.fn(),
      } as any;

      // Register tools in read-only mode
      registerGraphTools(readOnlyServer, mockGraphClient, true);

      // Check that update and delete tools are not registered
      const toolCalls = (readOnlyServer.tool as any).mock.calls.map((call: any) => call[0]);
      expect(toolCalls).not.toContain('update-planner-task');
      expect(toolCalls).not.toContain('delete-planner-task');

      // But read operations should still be available
      expect(toolCalls).toContain('list-planner-tasks');
      expect(toolCalls).toContain('get-planner-task');
    });
  });
});

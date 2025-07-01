import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { registerGraphTools } from '../src/graph-tools.js';
import GraphClient from '../src/graph-client.js';

// Mock GraphClient
vi.mock('../src/graph-client.js');

describe('Planner Plans Operations', () => {
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

  describe('list-user-planner-plans', () => {
    beforeEach(() => {
      // Register tools to set up the list-user-planner-plans tool
      registerGraphTools(server, mockGraphClient, false);
    });

    it('should be registered as a tool', () => {
      expect(server.tool).toHaveBeenCalledWith(
        'list-user-planner-plans',
        expect.any(String),
        expect.any(Object),
        expect.any(Object),
        expect.any(Function)
      );
    });

    it('should list all planner plans for the user', async () => {
      // Mock successful response with multiple plans
      mockGraphClient.graphRequest.mockResolvedValue({
        content: [
          {
            text: JSON.stringify({
              '@odata.count': 2,
              value: [
                {
                  id: 'plan-1',
                  title: 'IT Tasks',
                  createdDateTime: '2024-01-01T00:00:00Z',
                  owner: 'group-id-1',
                  '@odata.etag': 'W/"plan-etag-1"',
                },
                {
                  id: 'plan-2',
                  title: 'Marketing Tasks',
                  createdDateTime: '2024-01-02T00:00:00Z',
                  owner: 'group-id-2',
                  '@odata.etag': 'W/"plan-etag-2"',
                },
              ],
            }),
          },
        ],
        isError: false,
      });

      const listToolCall = (server.tool as any).mock.calls.find(
        (call: any) => call[0] === 'list-user-planner-plans'
      );
      expect(listToolCall).toBeDefined();

      const [, , , , handler] = listToolCall;

      const result = await handler({});

      expect(mockGraphClient.graphRequest).toHaveBeenCalledWith(
        '/me/planner/plans',
        expect.objectContaining({
          method: 'GET',
        })
      );

      expect(result.content).toHaveLength(1);
      expect(result.isError).toBeFalsy();

      const responseData = JSON.parse(result.content[0].text);
      expect(responseData.value).toHaveLength(2);
      expect(responseData.value[0].title).toBe('IT Tasks');
      expect(responseData.value[1].title).toBe('Marketing Tasks');
    });

    it('should support filtering plans by title', async () => {
      // Mock filtered response
      mockGraphClient.graphRequest.mockResolvedValue({
        content: [
          {
            text: JSON.stringify({
              value: [
                {
                  id: 'plan-1',
                  title: 'IT Tasks',
                  createdDateTime: '2024-01-01T00:00:00Z',
                  owner: 'group-id-1',
                },
              ],
            }),
          },
        ],
        isError: false,
      });

      const listToolCall = (server.tool as any).mock.calls.find(
        (call: any) => call[0] === 'list-user-planner-plans'
      );
      const [, , , , handler] = listToolCall;

      const result = await handler({
        filter: "contains(title,'IT')",
      });

      expect(mockGraphClient.graphRequest).toHaveBeenCalledWith(
        "/me/planner/plans?%24filter=contains(title%2C'IT')",
        expect.objectContaining({
          method: 'GET',
        })
      );

      expect(result.isError).toBeFalsy();
      const responseData = JSON.parse(result.content[0].text);
      expect(responseData.value).toHaveLength(1);
      expect(responseData.value[0].title).toBe('IT Tasks');
    });

    it('should support searching plans', async () => {
      mockGraphClient.graphRequest.mockResolvedValue({
        content: [
          {
            text: JSON.stringify({
              value: [
                {
                  id: 'plan-1',
                  title: 'IT Department Tasks',
                  createdDateTime: '2024-01-01T00:00:00Z',
                },
              ],
            }),
          },
        ],
        isError: false,
      });

      const listToolCall = (server.tool as any).mock.calls.find(
        (call: any) => call[0] === 'list-user-planner-plans'
      );
      const [, , , , handler] = listToolCall;

      const result = await handler({
        search: 'IT department',
      });

      expect(mockGraphClient.graphRequest).toHaveBeenCalledWith(
        '/me/planner/plans?%24search=IT%20department',
        expect.objectContaining({
          method: 'GET',
        })
      );

      expect(result.isError).toBeFalsy();
    });

    it('should support pagination with top and skip parameters', async () => {
      mockGraphClient.graphRequest.mockResolvedValue({
        content: [
          {
            text: JSON.stringify({
              '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/planner/plans?$skip=5',
              value: [
                {
                  id: 'plan-1',
                  title: 'Plan 1',
                },
              ],
            }),
          },
        ],
        isError: false,
      });

      const listToolCall = (server.tool as any).mock.calls.find(
        (call: any) => call[0] === 'list-user-planner-plans'
      );
      const [, , , , handler] = listToolCall;

      const result = await handler({
        top: 5,
        skip: 0,
      });

      expect(mockGraphClient.graphRequest).toHaveBeenCalledWith(
        '/me/planner/plans?%24top=5&%24skip=0',
        expect.objectContaining({
          method: 'GET',
        })
      );

      expect(result.isError).toBeFalsy();
    });

    it('should support selecting specific properties', async () => {
      mockGraphClient.graphRequest.mockResolvedValue({
        content: [
          {
            text: JSON.stringify({
              value: [
                {
                  id: 'plan-1',
                  title: 'IT Tasks',
                  // Only selected properties returned
                },
              ],
            }),
          },
        ],
        isError: false,
      });

      const listToolCall = (server.tool as any).mock.calls.find(
        (call: any) => call[0] === 'list-user-planner-plans'
      );
      const [, , , , handler] = listToolCall;

      const result = await handler({
        select: ['id', 'title'],
      });

      expect(mockGraphClient.graphRequest).toHaveBeenCalledWith(
        '/me/planner/plans?%24select=id%2Ctitle',
        expect.objectContaining({
          method: 'GET',
        })
      );

      expect(result.isError).toBeFalsy();
    });

    it('should support ordering plans', async () => {
      mockGraphClient.graphRequest.mockResolvedValue({
        content: [
          {
            text: JSON.stringify({
              value: [
                {
                  id: 'plan-2',
                  title: 'Marketing Tasks',
                  createdDateTime: '2024-01-02T00:00:00Z',
                },
                {
                  id: 'plan-1',
                  title: 'IT Tasks',
                  createdDateTime: '2024-01-01T00:00:00Z',
                },
              ],
            }),
          },
        ],
        isError: false,
      });

      const listToolCall = (server.tool as any).mock.calls.find(
        (call: any) => call[0] === 'list-user-planner-plans'
      );
      const [, , , , handler] = listToolCall;

      const result = await handler({
        orderby: ['createdDateTime desc'],
      });

      expect(mockGraphClient.graphRequest).toHaveBeenCalledWith(
        '/me/planner/plans?%24orderby=createdDateTime%20desc',
        expect.objectContaining({
          method: 'GET',
        })
      );

      expect(result.isError).toBeFalsy();
    });

    it('should handle empty results', async () => {
      mockGraphClient.graphRequest.mockResolvedValue({
        content: [
          {
            text: JSON.stringify({
              '@odata.count': 0,
              value: [],
            }),
          },
        ],
        isError: false,
      });

      const listToolCall = (server.tool as any).mock.calls.find(
        (call: any) => call[0] === 'list-user-planner-plans'
      );
      const [, , , , handler] = listToolCall;

      const result = await handler({});

      expect(result.isError).toBeFalsy();
      const responseData = JSON.parse(result.content[0].text);
      expect(responseData.value).toHaveLength(0);
      expect(responseData['@odata.count']).toBe(0);
    });

    it('should handle API errors gracefully', async () => {
      mockGraphClient.graphRequest.mockResolvedValue({
        content: [{ text: JSON.stringify({ error: 'Forbidden' }) }],
        isError: true,
      });

      const listToolCall = (server.tool as any).mock.calls.find(
        (call: any) => call[0] === 'list-user-planner-plans'
      );
      const [, , , , handler] = listToolCall;

      const result = await handler({});

      expect(result.isError).toBeTruthy();
      expect(result.content[0].text).toContain('Forbidden');
    });

    it('should combine multiple query parameters correctly', async () => {
      mockGraphClient.graphRequest.mockResolvedValue({
        content: [
          {
            text: JSON.stringify({
              value: [
                {
                  id: 'plan-1',
                  title: 'IT Tasks',
                },
              ],
            }),
          },
        ],
        isError: false,
      });

      const listToolCall = (server.tool as any).mock.calls.find(
        (call: any) => call[0] === 'list-user-planner-plans'
      );
      const [, , , , handler] = listToolCall;

      const result = await handler({
        filter: "contains(title,'IT')",
        top: 10,
        select: ['id', 'title'],
        orderby: ['title'],
      });

      expect(mockGraphClient.graphRequest).toHaveBeenCalledWith(
        "/me/planner/plans?%24filter=contains(title%2C'IT')&%24top=10&%24select=id%2Ctitle&%24orderby=title",
        expect.objectContaining({
          method: 'GET',
        })
      );

      expect(result.isError).toBeFalsy();
    });
  });

  describe('integration with existing plan tasks workflow', () => {
    beforeEach(() => {
      registerGraphTools(server, mockGraphClient, false);
    });

    it('should enable workflow: find plan by name → get tasks from plan', async () => {
      // Step 1: Mock plan search response
      mockGraphClient.graphRequest.mockResolvedValueOnce({
        content: [
          {
            text: JSON.stringify({
              value: [
                {
                  id: 'plan-123',
                  title: 'IT Tasks',
                  owner: 'group-id-1',
                },
              ],
            }),
          },
        ],
        isError: false,
      });

      // Step 2: Mock tasks from plan response
      mockGraphClient.graphRequest.mockResolvedValueOnce({
        content: [
          {
            text: JSON.stringify({
              value: [
                {
                  id: 'task-1',
                  title: 'Setup server',
                  percentComplete: 50,
                  planId: 'plan-123',
                },
                {
                  id: 'task-2',
                  title: 'Configure firewall',
                  percentComplete: 0,
                  planId: 'plan-123',
                },
              ],
            }),
          },
        ],
        isError: false,
      });

      // Find the tool handlers
      const listPlansCall = (server.tool as any).mock.calls.find(
        (call: any) => call[0] === 'list-user-planner-plans'
      );
      const listTasksCall = (server.tool as any).mock.calls.find(
        (call: any) => call[0] === 'list-plan-tasks'
      );

      const [, , , , listPlansHandler] = listPlansCall;
      const [, , , , listTasksHandler] = listTasksCall;

      // Step 1: Search for plan by name
      const planResult = await listPlansHandler({
        filter: "contains(title,'IT Tasks')",
      });

      expect(planResult.isError).toBeFalsy();
      const planData = JSON.parse(planResult.content[0].text);
      const foundPlan = planData.value[0];
      expect(foundPlan.title).toBe('IT Tasks');

      // Step 2: Get tasks from the found plan
      const tasksResult = await listTasksHandler({
        plannerPlanId: foundPlan.id,
      });

      expect(tasksResult.isError).toBeFalsy();
      const tasksData = JSON.parse(tasksResult.content[0].text);
      expect(tasksData.value).toHaveLength(2);
      expect(tasksData.value[0].title).toBe('Setup server');
      expect(tasksData.value[1].title).toBe('Configure firewall');
    });
  });

  describe('read-only mode', () => {
    it('should allow plan listing in read-only mode', () => {
      const readOnlyServer = {
        tool: vi.fn(),
      } as any;

      // Register tools in read-only mode
      registerGraphTools(readOnlyServer, mockGraphClient, true);

      // Check that plan listing tool is registered (it's a GET operation)
      const toolCalls = (readOnlyServer.tool as any).mock.calls.map((call: any) => call[0]);
      expect(toolCalls).toContain('list-user-planner-plans');
    });
  });
});

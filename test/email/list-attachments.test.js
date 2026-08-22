const handleListAttachments = require('../../email/list-attachments');
const { callGraphAPI } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleListAttachments', () => {
  const mockAccessToken = 'dummy_access_token';

  beforeEach(() => {
    callGraphAPI.mockClear();
    ensureAuthenticated.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
  });

  afterEach(() => {
    console.error.mockRestore();
  });

  test('requires an email id', async () => {
    const result = await handleListAttachments({});
    expect(result.content[0].text).toBe('Email ID is required.');
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('lists attachments with size and content type', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({
      value: [
        { id: 'att-1', name: 'Apolice.pdf', size: 175104, contentType: 'application/pdf', isInline: false }
      ]
    });

    const result = await handleListAttachments({ id: 'msg-1' });

    expect(callGraphAPI).toHaveBeenCalledWith(
      mockAccessToken,
      'GET',
      'me/messages/msg-1/attachments',
      null,
      { $select: 'id,name,size,contentType,isInline' }
    );
    expect(result.content[0].text).toContain('Found 1 attachment(s)');
    expect(result.content[0].text).toContain('Apolice.pdf');
    expect(result.content[0].text).toContain('id: att-1');
    expect(result.content[0].text).toContain('171 KB');
  });

  test('reports when there are no attachments', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({ value: [] });

    const result = await handleListAttachments({ id: 'msg-1' });

    expect(result.content[0].text).toBe('No attachments found on message msg-1.');
  });

  test('handles authentication error', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleListAttachments({ id: 'msg-1' });

    expect(result.content[0].text).toBe(
      "Authentication required. Please use the 'authenticate' tool first."
    );
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('handles Graph API error', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockRejectedValue(new Error('Graph API Error'));

    const result = await handleListAttachments({ id: 'msg-1' });

    expect(result.content[0].text).toBe('Error listing attachments: Graph API Error');
  });
});

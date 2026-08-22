const fs = require('fs');
const os = require('os');
const path = require('path');
const handleDownloadAttachment = require('../../email/download-attachment');
const { callGraphAPI, callGraphAPIRaw } = require('../../utils/graph-api');
const { ensureAuthenticated } = require('../../auth');

jest.mock('../../utils/graph-api');
jest.mock('../../auth');

describe('handleDownloadAttachment', () => {
  const mockAccessToken = 'dummy_access_token';
  let tmpDir;

  beforeEach(() => {
    callGraphAPI.mockClear();
    callGraphAPIRaw.mockClear();
    ensureAuthenticated.mockClear();
    jest.spyOn(console, 'error').mockImplementation(() => {});
    tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'outlook-mcp-attach-'));
  });

  afterEach(() => {
    console.error.mockRestore();
    fs.rmSync(tmpDir, { recursive: true, force: true });
  });

  test('requires an email id', async () => {
    const result = await handleDownloadAttachment({ destPath: tmpDir });
    expect(result.content[0].text).toBe('Email ID is required.');
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('requires a destPath', async () => {
    const result = await handleDownloadAttachment({ id: 'msg-1' });
    expect(result.content[0].text).toBe(
      'destPath is required: a directory to save into, or a full destination file path.'
    );
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('downloads the only file attachment into a directory', async () => {
    const contentBytes = Buffer.from('%PDF-1.4 fake policy content').toString('base64');
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({
      value: [
        {
          '@odata.type': '#microsoft.graph.fileAttachment',
          id: 'att-1',
          name: 'Apolice.pdf',
          size: 29,
          contentBytes
        }
      ]
    });

    const result = await handleDownloadAttachment({ id: 'msg-1', destPath: tmpDir });

    const savedPath = path.join(tmpDir, 'Apolice.pdf');
    expect(fs.existsSync(savedPath)).toBe(true);
    expect(fs.readFileSync(savedPath, 'utf8')).toBe('%PDF-1.4 fake policy content');
    expect(result.content[0].text).toContain('Saved attachment "Apolice.pdf"');
    expect(result.content[0].text).toContain(savedPath);
  });

  test('treats a non-existent extensionless path as a directory to create', async () => {
    const contentBytes = Buffer.from('content').toString('base64');
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({
      value: [
        { '@odata.type': '#microsoft.graph.fileAttachment', id: 'att-1', name: 'Apolice.pdf', size: 7, contentBytes }
      ]
    });

    // Regression test: a dest dir with no file extension and no trailing slash
    // (e.g. a brand-new folder) must NOT be treated as the literal file to write.
    const newDir = path.join(tmpDir, 'brand-new-folder');
    const result = await handleDownloadAttachment({ id: 'msg-1', destPath: newDir });

    const savedPath = path.join(newDir, 'Apolice.pdf');
    expect(fs.statSync(newDir).isDirectory()).toBe(true);
    expect(fs.readFileSync(savedPath, 'utf8')).toBe('content');
    expect(result.content[0].text).toContain(savedPath);
  });

  test('downloads to an explicit destination file path', async () => {
    const contentBytes = Buffer.from('content').toString('base64');
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({
      value: [
        { '@odata.type': '#microsoft.graph.fileAttachment', id: 'att-1', name: 'Apolice.pdf', size: 7, contentBytes }
      ]
    });

    const explicitPath = path.join(tmpDir, 'nested', 'renamed.pdf');
    const result = await handleDownloadAttachment({ id: 'msg-1', destPath: explicitPath });

    expect(fs.existsSync(explicitPath)).toBe(true);
    expect(fs.readFileSync(explicitPath, 'utf8')).toBe('content');
    expect(result.content[0].text).toContain(explicitPath);
  });

  test('selects attachment by attachmentId when there are multiple', async () => {
    const contentBytes = Buffer.from('second file').toString('base64');
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({
      value: [
        { '@odata.type': '#microsoft.graph.fileAttachment', id: 'att-1', name: 'a.pdf', size: 1, contentBytes: Buffer.from('a').toString('base64') },
        { '@odata.type': '#microsoft.graph.fileAttachment', id: 'att-2', name: 'b.pdf', size: 11, contentBytes }
      ]
    });

    const result = await handleDownloadAttachment({ id: 'msg-1', attachmentId: 'att-2', destPath: tmpDir });

    const savedPath = path.join(tmpDir, 'b.pdf');
    expect(fs.readFileSync(savedPath, 'utf8')).toBe('second file');
    expect(result.content[0].text).toContain('b.pdf');
  });

  test('selects attachment by name when there are multiple', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({
      value: [
        { '@odata.type': '#microsoft.graph.fileAttachment', id: 'att-1', name: 'a.pdf', size: 1, contentBytes: Buffer.from('a').toString('base64') },
        { '@odata.type': '#microsoft.graph.fileAttachment', id: 'att-2', name: 'b.pdf', size: 1, contentBytes: Buffer.from('b').toString('base64') }
      ]
    });

    const result = await handleDownloadAttachment({ id: 'msg-1', name: 'a.pdf', destPath: tmpDir });

    expect(fs.readFileSync(path.join(tmpDir, 'a.pdf'), 'utf8')).toBe('a');
    expect(result.content[0].text).toContain('a.pdf');
  });

  test('asks to disambiguate when multiple attachments and none specified', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({
      value: [
        { '@odata.type': '#microsoft.graph.fileAttachment', id: 'att-1', name: 'a.pdf', size: 1024, contentBytes: Buffer.from('a').toString('base64') },
        { '@odata.type': '#microsoft.graph.fileAttachment', id: 'att-2', name: 'b.pdf', size: 2048, contentBytes: Buffer.from('b').toString('base64') }
      ]
    });

    const result = await handleDownloadAttachment({ id: 'msg-1', destPath: tmpDir });

    expect(result.content[0].text).toContain('Message has 2 attachments');
    expect(result.content[0].text).toContain('a.pdf');
    expect(result.content[0].text).toContain('b.pdf');
    expect(fs.existsSync(path.join(tmpDir, 'a.pdf'))).toBe(false);
  });

  test('falls back to $value binary endpoint when contentBytes is missing', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({
      value: [
        { '@odata.type': '#microsoft.graph.fileAttachment', id: 'att-1', name: 'big.pdf', size: 5 * 1024 * 1024 }
      ]
    });
    callGraphAPIRaw.mockResolvedValue(Buffer.from('big binary content'));

    const result = await handleDownloadAttachment({ id: 'msg-1', destPath: tmpDir });

    expect(callGraphAPIRaw).toHaveBeenCalledWith(
      mockAccessToken,
      'me/messages/msg-1/attachments/att-1/$value'
    );
    expect(fs.readFileSync(path.join(tmpDir, 'big.pdf'), 'utf8')).toBe('big binary content');
    expect(result.content[0].text).toContain('big.pdf');
  });

  test('reports when message has no file attachments', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({
      value: [
        { '@odata.type': '#microsoft.graph.itemAttachment', id: 'att-1', name: 'Embedded event' }
      ]
    });

    const result = await handleDownloadAttachment({ id: 'msg-1', destPath: tmpDir });

    expect(result.content[0].text).toBe('No file attachments found on message msg-1.');
  });

  test('reports unknown attachmentId', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockResolvedValue({
      value: [
        { '@odata.type': '#microsoft.graph.fileAttachment', id: 'att-1', name: 'a.pdf', size: 1, contentBytes: Buffer.from('a').toString('base64') }
      ]
    });

    const result = await handleDownloadAttachment({ id: 'msg-1', attachmentId: 'does-not-exist', destPath: tmpDir });

    expect(result.content[0].text).toBe('Attachment ID does-not-exist not found on this message.');
  });

  test('handles authentication error', async () => {
    ensureAuthenticated.mockRejectedValue(new Error('Authentication required'));

    const result = await handleDownloadAttachment({ id: 'msg-1', destPath: tmpDir });

    expect(result.content[0].text).toBe(
      "Authentication required. Please use the 'authenticate' tool first."
    );
    expect(callGraphAPI).not.toHaveBeenCalled();
  });

  test('handles Graph API error', async () => {
    ensureAuthenticated.mockResolvedValue(mockAccessToken);
    callGraphAPI.mockRejectedValue(new Error('Graph API Error'));

    const result = await handleDownloadAttachment({ id: 'msg-1', destPath: tmpDir });

    expect(result.content[0].text).toBe('Error downloading attachment: Graph API Error');
  });
});

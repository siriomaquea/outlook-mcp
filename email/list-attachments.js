/**
 * List email attachments functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * List attachments handler
 * @param {object} args - Tool arguments
 * @param {string} args.id - Email ID (required)
 * @returns {object} - MCP response
 */
async function handleListAttachments(args) {
  const messageId = args.id;

  if (!messageId) {
    return {
      content: [{
        type: "text",
        text: "Email ID is required."
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const endpoint = `me/messages/${encodeURIComponent(messageId)}/attachments`;
    const queryParams = {
      $select: 'id,name,size,contentType,isInline'
    };

    const response = await callGraphAPI(accessToken, 'GET', endpoint, null, queryParams);
    const attachments = response.value || [];

    if (attachments.length === 0) {
      return {
        content: [{
          type: "text",
          text: `No attachments found on message ${messageId}.`
        }]
      };
    }

    const lines = attachments.map(a =>
      `- ${a.name} (id: ${a.id}, ${Math.round((a.size || 0) / 1024)} KB, ${a.contentType || 'unknown type'}${a.isInline ? ', inline' : ''})`
    );

    return {
      content: [{
        type: "text",
        text: `Found ${attachments.length} attachment(s) on message ${messageId}:\n${lines.join('\n')}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{
          type: "text",
          text: "Authentication required. Please use the 'authenticate' tool first."
        }]
      };
    }

    return {
      content: [{
        type: "text",
        text: `Error listing attachments: ${error.message}`
      }]
    };
  }
}

module.exports = handleListAttachments;

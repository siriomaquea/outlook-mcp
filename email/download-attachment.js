/**
 * Download email attachment functionality
 *
 * Saves a file attachment from an email to the local filesystem, since this
 * MCP server runs on the same machine as the calling client.
 */
const fs = require('fs');
const path = require('path');
const { callGraphAPI, callGraphAPIRaw } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Download attachment handler
 * @param {object} args - Tool arguments
 * @param {string} args.id - Email ID (required)
 * @param {string} args.attachmentId - Attachment ID (optional if the email has exactly one file attachment)
 * @param {string} args.name - Attachment file name, alternative to attachmentId
 * @param {string} args.destPath - Local path to save to: an existing/target directory, or a full file path (required)
 * @returns {object} - MCP response
 */
async function handleDownloadAttachment(args) {
  const messageId = args.id;
  const attachmentId = args.attachmentId;
  const attachmentName = args.name;
  const destPath = args.destPath;

  if (!messageId) {
    return {
      content: [{ type: "text", text: "Email ID is required." }]
    };
  }

  if (!destPath) {
    return {
      content: [{ type: "text", text: "destPath is required: a directory to save into, or a full destination file path." }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();

    const endpoint = `me/messages/${encodeURIComponent(messageId)}/attachments`;
    const response = await callGraphAPI(accessToken, 'GET', endpoint);
    const attachments = (response.value || []).filter(a => a['@odata.type'] === '#microsoft.graph.fileAttachment');

    if (attachments.length === 0) {
      return {
        content: [{ type: "text", text: `No file attachments found on message ${messageId}.` }]
      };
    }

    let target;
    if (attachmentId) {
      target = attachments.find(a => a.id === attachmentId);
      if (!target) {
        return {
          content: [{ type: "text", text: `Attachment ID ${attachmentId} not found on this message.` }]
        };
      }
    } else if (attachmentName) {
      target = attachments.find(a => a.name === attachmentName);
      if (!target) {
        return {
          content: [{
            type: "text",
            text: `No attachment named "${attachmentName}" found. Available: ${attachments.map(a => a.name).join(', ')}`
          }]
        };
      }
    } else if (attachments.length === 1) {
      target = attachments[0];
    } else {
      return {
        content: [{
          type: "text",
          text: `Message has ${attachments.length} attachments. Specify attachmentId or name:\n` +
            attachments.map(a => `- ${a.name} (id: ${a.id}, ${Math.round((a.size || 0) / 1024)} KB)`).join('\n')
        }]
      };
    }

    let contentBuffer;
    if (target.contentBytes) {
      contentBuffer = Buffer.from(target.contentBytes, 'base64');
    } else {
      // Large attachments may omit contentBytes from the /attachments listing;
      // fall back to the binary $value endpoint.
      contentBuffer = await callGraphAPIRaw(
        accessToken,
        `me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(target.id)}/$value`
      );
    }

    // A destPath is treated as a directory (attachment saved inside it under its
    // original name) when it ends with a separator, already exists as a directory,
    // or has no file extension. Otherwise it's treated as the exact destination file path.
    const looksLikeDir = destPath.endsWith('/') || destPath.endsWith('\\') ||
      (fs.existsSync(destPath) && fs.statSync(destPath).isDirectory()) ||
      path.extname(destPath) === '';

    let finalPath;
    if (looksLikeDir) {
      fs.mkdirSync(destPath, { recursive: true });
      finalPath = path.join(destPath, target.name);
    } else {
      finalPath = destPath;
      fs.mkdirSync(path.dirname(finalPath), { recursive: true });
    }

    fs.writeFileSync(finalPath, contentBuffer);

    return {
      content: [{
        type: "text",
        text: `Saved attachment "${target.name}" (${Math.round(contentBuffer.length / 1024)} KB) to ${finalPath}`
      }]
    };
  } catch (error) {
    if (error.message === 'Authentication required') {
      return {
        content: [{ type: "text", text: "Authentication required. Please use the 'authenticate' tool first." }]
      };
    }

    return {
      content: [{ type: "text", text: `Error downloading attachment: ${error.message}` }]
    };
  }
}

module.exports = handleDownloadAttachment;

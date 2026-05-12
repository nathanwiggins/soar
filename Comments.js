function getCommentsByTopic(topicId) {
  const normalizedTopicId = topicId ? topicId.toString().trim() : '';
  if (!normalizedTopicId) {
    return JSON.stringify({ success: false, error: 'Topic_ID is required.', comments: [] });
  }

  try {
    const comments = getTableData('Comments')
      .filter((comment) => (comment.Topic_ID || '').toString().trim() === normalizedTopicId)
      .map((comment) => {
        const normalized = {};
        Object.keys(comment).forEach((key) => {
          normalized[key] = normalizeValueForClient(comment[key]);
        });
        return normalized;
      });
    return JSON.stringify({ success: true, topicId: normalizedTopicId, comments });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to load comments.',
      comments: []
    });
  }
}
function extractMentionedUsers(content, users) {
  if (!content) return [];
  const mentionPattern = /@([a-zA-Z0-9._-]+)/g;
  const matches = content.matchAll(mentionPattern);
  const mentions = new Set();
  for (const match of matches) {
    if (match && match[1]) {
      mentions.add(match[1].toLowerCase());
    }
  }
  if (mentions.size === 0) return [];

  const toHandle = (user) => {
    const emailLocalPart = normalizeEmail(user && user.Email).split('@')[0];
    const fallbackHandle = (user && user.Name ? user.Name : '')
      .toString()
      .trim()
      .toLowerCase()
      .replace(/\s+/g, '.')
      .replace(/[^a-z0-9._-]/g, '');
    return emailLocalPart || fallbackHandle;
  };

  return users.filter((user) => {
    const handle = toHandle(user);
    return handle && mentions.has(handle);
  });
}
function addComment(topicId, commentInput) {
  try {
    const normalizedTopicId = ensureTaskExists(topicId);
    const content = commentInput && commentInput.content ? commentInput.content.toString().trim() : '';
    if (!content) {
      throw new Error('Comment content is required.');
    }

    const currentUserEmail = normalizeEmail(getCurrentUser());
    const currentUserId = getCurrentUserIdByEmail(currentUserEmail);
    if (!currentUserId) {
      throw new Error('Could not determine commenter_ID from current user email.');
    }

    const commentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Comments');
    if (!commentsSheet) {
      throw new Error('Comments sheet was not found.');
    }

    const headers = commentsSheet.getRange(1, 1, 1, commentsSheet.getLastColumn()).getValues()[0];
    const headerIndex = getHeaderIndex(headers);
    const now = new Date();
    const newRow = new Array(headers.length).fill('');

    if (headerIndex.Comment_ID !== undefined) newRow[headerIndex.Comment_ID] = generateNextId('Comments', 'C');
    if (headerIndex.Topic_ID !== undefined) newRow[headerIndex.Topic_ID] = normalizedTopicId;
    if (headerIndex.Commenter_ID !== undefined) newRow[headerIndex.Commenter_ID] = currentUserId;
    if (headerIndex.Timestamp !== undefined) newRow[headerIndex.Timestamp] = now;
    if (headerIndex.Content !== undefined) newRow[headerIndex.Content] = content;

    appendRows(commentsSheet, [newRow]);
    invalidateTableCache('Comments');

    const createdComment = {};
    headers.forEach((header, index) => {
      const value = newRow[index];
      createdComment[header] = value instanceof Date ? value.toISOString() : value;
    });

    const users = getTableData('Users');
    const commenter = users.find((user) => (user.User_ID || '').toString().trim() === currentUserId) || null;
    const mentionedUsers = extractMentionedUsers(content, users).filter(
      (user) => (user.User_ID || '').toString().trim() !== currentUserId
    );
    sendMentionNotifications(content, normalizedTopicId, commenter, mentionedUsers);

    return JSON.stringify({ success: true, comment: createdComment });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to add comment.'
    });
  }
}
function deleteComment(commentId) {
  const normalizedCommentId = commentId ? commentId.toString().trim() : '';
  if (!normalizedCommentId) {
    return JSON.stringify({ success: false, error: 'Comment ID is required.' });
  }

  const commentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Comments');
  if (!commentsSheet) {
    return JSON.stringify({ success: false, error: 'Comments sheet was not found.' });
  }

  try {
    const data = commentsSheet.getDataRange().getValues();
    if (data.length <= 1) {
      throw new Error('Comments sheet has no data rows.');
    }

    const headers = data[0];
    const commentIdColumnIndex = headers.indexOf('Comment_ID');
    if (commentIdColumnIndex === -1) {
      throw new Error('Comments sheet is missing Comment_ID column.');
    }

    const rowIndex = data.findIndex((row, index) => index > 0 && row[commentIdColumnIndex] === normalizedCommentId);
    if (rowIndex < 0) {
      throw new Error('Comment not found.');
    }

    deleteRowsBySheetIndexes(commentsSheet, [rowIndex + 1]);
    invalidateTableCache('Comments');
    return JSON.stringify({ success: true, commentId: normalizedCommentId });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to delete comment.'
    });
  }
}
function resolveComment(commentId) {
  const normalizedCommentId = commentId ? commentId.toString().trim() : '';
  if (!normalizedCommentId) {
    return JSON.stringify({ success: false, error: 'Comment ID is required.' });
  }

  const commentsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Comments');
  if (!commentsSheet) {
    return JSON.stringify({ success: false, error: 'Comments sheet was not found.' });
  }

  try {
    const data = commentsSheet.getDataRange().getValues();
    if (data.length <= 1) {
      throw new Error('Comments sheet has no data rows.');
    }

    const headers = data[0];
    const commentIdColumnIndex = headers.indexOf('Comment_ID');
    const topicIdColumnIndex = headers.indexOf('Topic_ID');
    const commenterIdColumnIndex = headers.indexOf('Commenter_ID');
    const contentColumnIndex = headers.indexOf('Content');

    if (commentIdColumnIndex === -1) throw new Error('Comments sheet is missing Comment_ID column.');
    if (commenterIdColumnIndex === -1) throw new Error('Comments sheet is missing Commenter_ID column.');

    const rowIndex = data.findIndex((row, index) => index > 0 && row[commentIdColumnIndex] === normalizedCommentId);
    if (rowIndex < 0) {
      throw new Error('Comment not found.');
    }

    const row = data[rowIndex];
    const commenterId = commenterIdColumnIndex > -1 ? row[commenterIdColumnIndex] : '';
    const commenter = getUserById(commenterId);
    const topicId = topicIdColumnIndex > -1 ? row[topicIdColumnIndex] : '';
    const content = contentColumnIndex > -1 ? row[contentColumnIndex] : '';

    deleteRowsBySheetIndexes(commentsSheet, [rowIndex + 1]);
    invalidateTableCache('Comments');

    const commenterEmail = commenter ? normalizeEmail(commenter.Email) : '';
    if (commenterEmail && isNotificationEnabledForUser(commenter, 'commentsAndMentions')) {
      const resolverEmail = normalizeEmail(getCurrentUser());
      const resolverUserId = getCurrentUserIdByEmail(resolverEmail);
      const resolver = getUserById(resolverUserId);
      const resolverName = resolver && resolver.Name ? resolver.Name : 'A teammate';
      const details = getTaskNotificationDetails(getTaskById(topicId));
      const subject = `Your comment was resolved on ${details.taskTitle}`;
      const body = `${resolverName} resolved your comment.

${formatTaskNotificationBody(details)}

Resolved comment:
${content}
`;
      safeSendEmail(commenterEmail, subject, body);
    }

    return JSON.stringify({ success: true, commentId: normalizedCommentId, resolved: true });
  } catch (error) {
    return JSON.stringify({
      success: false,
      error: error && error.message ? error.message : 'Failed to resolve comment.'
    });
  }
}

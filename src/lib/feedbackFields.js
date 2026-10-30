export const FEEDBACK_STATUSES = ['New', 'In progress', 'Resolved', 'Closed'];
export const FEEDBACK_STATUS_XML = '<Field Type="Choice" DisplayName="Feedback status" Name="FeedbackStatus" StaticName="FeedbackStatus" Required="FALSE"><Default>New</Default><CHOICES>' + FEEDBACK_STATUSES.map(value => `<CHOICE>${value}</CHOICE>`).join('') + '</CHOICES></Field>';

// Upgrade existing lists only when an owner saves an update. No ACL or item changes.
export async function ensureFeedbackStatus(store, path) {
  const read = async () => {
    // A collection query returns [] for an absent column; direct field lookups
    // can report a missing field as 400, indistinguishable from other errors.
    const body = await store.get(`${path}/fields?$filter=InternalName eq 'FeedbackStatus'`);
    const fields = body.value || body.d?.results;
    if (!Array.isArray(fields)) throw new Error('Could not verify the feedback status column.');
    return fields.find(field => field.InternalName === 'FeedbackStatus');
  };
  let field = await read();
  if (!field) {
    try {
      await store.post(`${path}/fields/createfieldasxml`, { verbose: true, body: { parameters: { __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' }, SchemaXml: FEEDBACK_STATUS_XML, Options: 24 } } });
    } catch (error) {
      // Another owner may have added the field concurrently. Verify before proceeding.
      field = await read();
      if (!field) throw error;
    }
    field ||= await read();
  }
  const choices = field?.Choices?.results || field?.Choices;
  if (field?.TypeAsString !== 'Choice' || !Array.isArray(choices) || !FEEDBACK_STATUSES.every(value => choices.includes(value))) throw new Error('Feedback status could not be configured. Ask a site owner with Manage Lists to review the FeedbackStatus choice column.');
}

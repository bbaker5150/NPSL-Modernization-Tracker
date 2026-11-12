const xml = value => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');

// CSOM executes AddItem + field setters + Update without loading the created
// item. REST POST /items serializes that item back to an add-only caller, which
// can fail after creation. Never follow an uncertain write with another create.
// Protocol types: SP.RequestContext and SP.ListItemCreationInformation.
export function feedbackCreateRequest(title, report) {
  const fields = { Title: report.title, FeedbackKind: report.kind, FeedbackDetails: report.details };
  if (report.attachments?.length) fields.FeedbackAttachments = JSON.stringify(report.attachments);
  return `<Request xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0" ApplicationName="Modernization Tracker">
<Actions>${Object.entries(fields).map(([name, value], index) => `<Method Name="SetFieldValue" Id="${10 + index}" ObjectPathId="4"><Parameters><Parameter Type="String">${name}</Parameter><Parameter Type="String">${xml(value)}</Parameter></Parameters></Method>`).join('')}<Method Name="Update" Id="20" ObjectPathId="4" /></Actions>
<ObjectPaths><StaticProperty Id="0" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" /><Property Id="1" ParentId="0" Name="Web" /><Property Id="2" ParentId="1" Name="Lists" /><Method Id="3" ParentId="2" Name="GetByTitle"><Parameters><Parameter Type="String">${xml(title)}</Parameter></Parameters></Method><Method Id="4" ParentId="3" Name="AddItem"><Parameters><Parameter TypeId="{54cdbee5-0897-44ac-829f-411557fa11be}"><Property Name="FolderUrl" Type="Null" /><Property Name="LeafName" Type="Null" /><Property Name="UnderlyingObjectType" Type="Number">0</Property></Parameter></Parameters></Method></ObjectPaths></Request>`;
}

const unconfirmed = message => Object.assign(new Error(message), { submissionUncertain: true });

export async function submitPrivateFeedback(store, titles, report) {
  for (const title of titles) {
    let response;
    try { response = await store.post('/_vti_bin/client.svc/ProcessQuery', { raw: true, headers: { 'Content-Type': 'text/xml; charset=utf-8', Accept: 'application/json' }, body: feedbackCreateRequest(title, report) }); }
    catch (cause) { throw unconfirmed(`SharePoint did not confirm your submission. Ask a site owner to check the inbox before you retry. ${cause.message}`); }
    const header = Array.isArray(response) && response[0];
    if (!header || !Object.hasOwn(header, 'ErrorInfo')) throw unconfirmed('SharePoint returned no submission confirmation. Ask a site owner to check the inbox before retrying.');
    if (header.ErrorInfo === null) return;
    const error = header.ErrorInfo;
    // Only this explicit list-resolution failure is known to occur before
    // AddItem. A generic "Item does not exist", HTTP 500, or timeout is ambiguous.
    if (error?.ErrorTypeName === 'System.ArgumentException' && error.ErrorCode === -2147024809 && String(error.ErrorMessage).startsWith(`List '${title}' does not exist at site with URL`)) continue;
    throw unconfirmed(`SharePoint could not confirm your submission: ${error?.ErrorMessage || 'Unknown server response'}. Ask a site owner to check the inbox before retrying.`);
  }
  throw new Error('Feedback storage is missing or unavailable. A site owner must open feedback and complete Set up private feedback.');
}

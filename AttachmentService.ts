import { IAttachment, IAttachmentServiceContext } from './types/IAttachment';
import { IAttachmentService } from './types/IAttachmentService';

const FIELDS = 'annotationid,subject,notetext,filename,filesize,mimetype,documentbody,createdon,_createdby_value,isdocument,_objectid_value,objecttypecode';

export class AttachmentService implements IAttachmentService {
  async getAttachments(context: IAttachmentServiceContext): Promise<IAttachment[]> {
    try {
      const query = `?$select=${FIELDS}&$filter=_objectid_value eq '${context.recordId}' and isdocument eq true&$orderby=createdon desc`;
      const response = await context.webAPI.retrieveMultipleRecords('annotation', query);

      return (response.entities as Record<string, unknown>[]).map(entity => ({
        annotationId: entity.annotationid as string,
        subject: (entity.subject as string) ?? (entity.filename as string) ?? 'Untitled',
        noteText: (entity.notetext as string) ?? '',
        fileName: (entity.filename as string) ?? '',
        fileSize: (entity.filesize as number) ?? 0,
        mimeType: (entity.mimetype as string) ?? '',
        documentBody: (entity.documentbody as string) ?? '',
        createdOn: new Date(entity.createdon as string),
        createdBy: (entity['_createdby_value@OData.Community.Display.V1.FormattedValue'] as string) ?? 'Unknown',
        isDocument: entity.isdocument as boolean,
        objectId: entity._objectid_value as string,
        objectTypeCode: Number(entity.objecttypecode ?? 0)
      }));
    } catch (error) {
      console.error('❌ Error retrieving attachments:', error);
      throw new Error('Failed to retrieve attachments.');
    }
  }

  async uploadAttachment(
    context: IAttachmentServiceContext,
    file: File,
    subject?: string
  ): Promise<IAttachment> {
    console.log("📌 Upload context:", context);
    try {
      const base64Data = await this.convertFileToBase64(file);
      const entityBindPath = `/${context.entityName}s(${context.recordId})`;

      const annotationData: Record<string, unknown> = {
        subject: subject ?? file.name,
        notetext: `Uploaded file: ${file.name}`,
        filename: file.name,
        mimetype: file.type,
        documentbody: base64Data,
        isdocument: true,
        [`objectid_${context.entityName}@odata.bind`]: entityBindPath
      };

      const response = await context.webAPI.createRecord('annotation', annotationData);
      return await this.getAttachmentById(context, response.id);
    } catch (error) {
      console.error('❌ Error uploading attachment:', error);
      throw new Error('Failed to upload attachment.');
    }
  }

  async deleteAttachment(context: IAttachmentServiceContext, attachmentId: string): Promise<void> {
    try {
      await context.webAPI.deleteRecord('annotation', attachmentId);
    } catch (error) {
      console.error('❌ Error deleting attachment:', error);
      throw new Error('Failed to delete attachment.');
    }
  }

downloadAttachment(attachment: IAttachment): void {
  try {
    const byteCharacters = atob(attachment.documentBody);
    const byteNumbers: number[] = Array.from({ length: byteCharacters.length }, (_, i) =>
                      byteCharacters.charCodeAt(i));

    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], { type: attachment.mimeType });

    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = attachment.fileName;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  } catch (error) {
    console.error('❌ Error downloading attachment:', error);
    throw new Error('Failed to download attachment.');
  }
}

  private async getAttachmentById(
    context: IAttachmentServiceContext,
    attachmentId: string
  ): Promise<IAttachment> {
    try {
      const response = (await context.webAPI.retrieveRecord('annotation', attachmentId, `?$select=${FIELDS}`)) as Record<string, unknown>;
      return {
        annotationId: response.annotationid as string,
        subject: (response.subject as string) ?? (response.filename as string) ?? 'Untitled',
        noteText: (response.notetext as string) ?? '',
        fileName: (response.filename as string) ?? '',
        fileSize: (response.filesize as number) ?? 0,
        mimeType: (response.mimetype as string) ?? '',
        documentBody: (response.documentbody as string) ?? '',
        createdOn: new Date(response.createdon as string),
        createdBy: (response['_createdby_value@OData.Community.Display.V1.FormattedValue'] as string) ?? 'Unknown',
        isDocument: response.isdocument as boolean,
        objectId: response._objectid_value as string,
        objectTypeCode: Number(response.objecttypecode ?? 0)
      };
    } catch (error) {
      console.error('❌ Error retrieving single attachment by ID:', error);
      throw new Error('Failed to retrieve uploaded attachment.');
    }
  }

  private convertFileToBase64(file: File): Promise<string> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        const result = reader.result as string;
        const base64String = result?.split(',')[1] ?? '';
        resolve(base64String);
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }
}
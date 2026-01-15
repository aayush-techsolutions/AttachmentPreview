export interface IAttachment {
  annotationId: string;
  subject: string;
  noteText: string;
  fileName: string;
  fileSize: number;
  mimeType: string;
  documentBody: string;
  createdOn: Date;
  createdBy: string;
  isDocument: boolean;
  objectId: string;
  objectTypeCode: number;
}

export interface IAttachmentServiceContext {
  webAPI: ComponentFramework.WebApi;
  recordId: string;
  entityName: string;
}
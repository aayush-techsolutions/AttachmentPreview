import { IAttachment, IAttachmentServiceContext } from './IAttachment';

export interface IAttachmentService {
  getAttachments(context: IAttachmentServiceContext): Promise<IAttachment[]>;
  uploadAttachment(
    context: IAttachmentServiceContext,
    file: File,
    subject?: string
  ): Promise<IAttachment>;
  
  deleteAttachment(
    context: IAttachmentServiceContext,
    attachmentId: string
  ): Promise<void>;
  downloadAttachment(attachment: IAttachment): void;
}
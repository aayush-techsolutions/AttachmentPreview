import { IInputs, IOutputs } from './generated/ManifestTypes';
import * as React from 'react';
import { createRoot, Root } from 'react-dom/client';
import { AttachmentPreviewComponent, IAttachmentPreviewProps } from './App';
import { AttachmentService } from './AttachmentService';
import { IAttachment, IAttachmentServiceContext } from './types/IAttachment';

export class AttachmentPreviewControl implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private container!: HTMLDivElement;
  private context!: ComponentFramework.Context<IInputs>;
  private attachmentService: AttachmentService;
  private attachments: IAttachment[] = [];
  private isLoading = false;
  private notifyOutputChanged!: () => void;
  private root: Root | null = null;

  constructor() {
    this.attachmentService = new AttachmentService();
  }

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this.context = context;
    this.container = container;
    this.notifyOutputChanged = notifyOutputChanged;
    void this.loadAttachments(); // async fire-and-forget
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this.context = context;
    this.renderComponent();
  }

  public getOutputs(): IOutputs {
    return {};
  }

  public destroy(): void {
    this.root?.unmount();
    this.root = null;
  }

  private async loadAttachments(): Promise<void> {
    try {
      this.isLoading = true;
      const serviceContext = this.createServiceContext();
      console.log('📦 Loading attachments with context:', serviceContext);

      if (!serviceContext.recordId || !serviceContext.entityName) {
        console.warn('❌ Missing recordId or entityName, skipping attachment fetch.');
        this.attachments = [];
        return;
      }

      this.attachments = await this.attachmentService.getAttachments(serviceContext);
      console.log('📎 Attachments retrieved:', this.attachments);
    } catch (error) {
      console.error('❌ Failed to load attachments:', error);
      this.attachments = [];
    } finally {
      this.isLoading = false;
      this.renderComponent();
    }
  }

private createServiceContext(): IAttachmentServiceContext {
  const modeWithContextInfo = this.context.mode as unknown as {
    contextInfo?: {
      entityId: string;
      entityTypeName: string;
    };
  };
  console.log("📌 context.mode.contextInfo:", modeWithContextInfo.contextInfo);
  const entityId = modeWithContextInfo.contextInfo?.entityId;
  const entityName = modeWithContextInfo.contextInfo?.entityTypeName;

  if (!entityId || !entityName) {
    console.warn('⚠️ Entity reference is not available.');
    return {
      webAPI: this.context.webAPI,
      recordId: '',
      entityName: ''
    };
  }

  return {
    webAPI: this.context.webAPI,
    recordId: entityId,
    entityName: entityName
  };
}

  private handleUpload = async (files: File[]): Promise<void> => {
    const serviceContext = this.createServiceContext();
    console.log('Upload context:', serviceContext);

    if (!serviceContext.recordId || !serviceContext.entityName) {
      console.warn('❌ Cannot upload without valid recordId and entityName.');
      return;
    }

    try {
      for (const file of files) {
        const newAttachment = await this.attachmentService.uploadAttachment(serviceContext, file);
        this.attachments.unshift(newAttachment);
      }
    } catch (error) {
      console.error('❌ Failed to upload files:', error);
    }

    this.renderComponent();
  };

  private handleDelete = async (attachmentId: string): Promise<void> => {
    const serviceContext = this.createServiceContext();

    if (!serviceContext.recordId || !serviceContext.entityName) {
      console.warn('❌ Cannot delete without valid recordId and entityName.');
      return;
    }

    try {
      await this.attachmentService.deleteAttachment(serviceContext, attachmentId);
      this.attachments = this.attachments.filter(att => att.annotationId !== attachmentId);
    } catch (error) {
      console.error('❌ Failed to delete attachment:', error);
    }

    this.renderComponent();
  };

  private handlePreview = (attachment: IAttachment): void => {
    void this.attachmentService.downloadAttachment(attachment);
  };

  private renderComponent(): void {
    const props: IAttachmentPreviewProps = {
      attachments: this.attachments,
      onUpload: this.handleUpload,
      onDelete: this.handleDelete,
      onPreview: this.handlePreview,
      maxFileSize: this.context.parameters.maxFileSize?.raw ?? 10,
      allowedFileTypes: (this.context.parameters.allowedFileTypes?.raw ?? 'jpg,jpeg,png,gif,pdf,docx,xlsx,pptx,txt,mp4').split(','),
      showPreview: this.context.parameters.showPreview?.raw ?? true,
      isLoading: this.isLoading
    };
    window.refreshPCFAttachments = this.loadAttachments.bind(this); // for internal refresh

    this.root ??= createRoot(this.container);
    this.root.render(React.createElement(AttachmentPreviewComponent, props));
  }
}
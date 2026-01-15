// (unchanged imports)
import * as React from 'react';
import { useState, useEffect, useCallback, useRef } from 'react';
import { IAttachment } from './types/IAttachment';
const BYTES_IN_MB = 1024 * 1024;

export interface IAttachmentPreviewProps {
  attachments: IAttachment[];
  onUpload: (files: File[]) => Promise<void>;
  onDelete: (attachmentId: string) => Promise<void>;
  onPreview: (attachment: IAttachment) => void;
  maxFileSize: number;
  allowedFileTypes: string[];
  showPreview: boolean;
  isLoading: boolean;
}

import {
  DetailsList, DetailsListLayoutMode, IColumn,
  CommandBar, ICommandBarItemProps,
  MessageBar, MessageBarType,
  Spinner, SpinnerSize,
  Panel, PanelType,
  DefaultButton, PrimaryButton,
  Stack, Text,
  IconButton, Image,
  Dialog, DialogType, DialogFooter
} from '@fluentui/react';

export const AttachmentPreviewComponent: React.FC<IAttachmentPreviewProps> = ({
  attachments,
  onUpload,
  onDelete,
  onPreview,
  maxFileSize,
  allowedFileTypes,
  showPreview,
  isLoading
}) => {
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [selectedAttachment, setSelectedAttachment] = useState<IAttachment | null>(null);
  const [showPanel, setShowPanel] = useState(false);
  const [showDeleteDialog, setShowDeleteDialog] = useState(false);
  const [attachmentToDelete, setAttachmentToDelete] = useState<IAttachment | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const [message, setMessage] = useState<{ text: string; type: MessageBarType } | null>(null);

  const allowedTypesArray = allowedFileTypes.map(type => type.trim().toLowerCase());

  const handleRefreshClick = () => {
    if (typeof window.refreshPCFAttachments === 'function') {
      void window.refreshPCFAttachments();
    } else {
      setMessage({ text: 'Refresh not available', type: MessageBarType.warning });
    }
  };

  const columns: IColumn[] = [
    {
      key: 'fileName',
      name: 'File Name',
      fieldName: 'fileName',
      minWidth: 200,
      maxWidth: 300,
      isResizable: true,
      onRender: (item: IAttachment) => (
        <Stack horizontal verticalAlign="center">
          <IconButton
            iconProps={{ iconName: getFileIcon(item.mimeType) }}
            title={item.mimeType}
            styles={{ root: { marginRight: 8 } }}
          />
          <Text>{item.fileName}</Text>
        </Stack>
      )
    },
    {
      key: 'fileSize',
      name: 'Size',
      fieldName: 'fileSize',
      minWidth: 80,
      maxWidth: 100,
      isResizable: true,
      onRender: (item: IAttachment) => <Text>{formatFileSize(item.fileSize)}</Text>
    },
    {
      key: 'createdOn',
      name: 'Created',
      fieldName: 'createdOn',
      minWidth: 120,
      maxWidth: 150,
      isResizable: true,
      onRender: (item: IAttachment) => <Text>{new Date(item.createdOn).toLocaleDateString()}</Text>
    },
    {
      key: 'createdBy',
      name: 'Created By',
      fieldName: 'createdBy',
      minWidth: 120,
      maxWidth: 150,
      isResizable: true
    },
    {
      key: 'actions',
      name: 'Actions',
      minWidth: 100,
      maxWidth: 120,
      onRender: (item: IAttachment) => (
        <Stack horizontal>
          {showPreview && (
            <IconButton iconProps={{ iconName: 'View' }} title="Preview" onClick={() => handlePreview(item)} />
          )}
          <IconButton iconProps={{ iconName: 'Download' }} title="Download" onClick={() => onPreview(item)} />
          <IconButton iconProps={{ iconName: 'Delete' }} title="Delete" onClick={() => handleDeleteClick(item)} />
        </Stack>
      )
    }
  ];

  const handleUploadClick = useCallback(() => {
    const input = document.createElement('input');
    input.type = 'file';
    input.multiple = true;
    input.accept = allowedTypesArray.map(type => `.${type}`).join(',');
    input.onchange = async (e: Event) => {
      const target = e.target as HTMLInputElement;
      if (target.files) {
        await handleFileUpload(Array.from(target.files));
      }
    };
    input.click();
  }, [allowedTypesArray]);

  const commandBarItems: ICommandBarItemProps[] = [
    {
      key: 'upload',
      text: 'Upload Files',
      iconProps: { iconName: 'Upload' },
      onClick: handleUploadClick
    },
    {
      key: 'refresh',
      text: 'Refresh',
      iconProps: { iconName: 'Refresh' },
      onClick: handleRefreshClick
    }
  ];

  const refreshAttachments = () => {
  try {
    setMessage(null); // clear existing messages
    setSelectedAttachment(null); // reset preview
    setShowPanel(false); // close preview panel
    setShowDeleteDialog(false); // close delete dialog
    handleRefreshClick();
  } catch (err) {
    setMessage({ text: 'Failed to refresh attachments', type: MessageBarType.error });
  }
};

  const handleFileUpload = async (files: File[]) => {
    const validFiles = files.filter(file => {
      const extension = file.name.split('.').pop()?.toLowerCase();
      const isValidType = allowedTypesArray.includes(extension ?? '');
      const isValidSize = file.size <= maxFileSize * BYTES_IN_MB;

      if (!isValidType) {
        setMessage({ text: `File type .${extension} is not allowed`, type: MessageBarType.error });
        return false;
      }
      if (!isValidSize) {
        setMessage({ text: `File ${file.name} exceeds max size limit`, type: MessageBarType.error });
        return false;
      }
      return true;
    });

    if (validFiles.length > 0) {
      try {
        await onUpload(validFiles);
        setMessage({ text: `Uploaded ${validFiles.length} file(s) successfully`, type: MessageBarType.success });
      } catch (error: unknown) {
        setMessage({ text: 'Upload failed', type: MessageBarType.error });
      }
    }
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setDragOver(false);
    const files = Array.from(e.dataTransfer.files);
    void handleFileUpload(files);
  };

  const handlePreview = (attachment: IAttachment) => {
    setSelectedAttachment(attachment);
    setShowPanel(true);
  };

  const handleDeleteClick = (attachment: IAttachment) => {
    setAttachmentToDelete(attachment);
    setShowDeleteDialog(true);
  };

  const handleDeleteConfirm = async (): Promise<void> => {
    if (attachmentToDelete) {
      try {
        await onDelete(attachmentToDelete.annotationId);
        setMessage({ text: 'Deleted successfully', type: MessageBarType.success });
      } catch (error: unknown) {
        setMessage({ text: 'Delete failed', type: MessageBarType.error });
      }
    }
    setShowDeleteDialog(false);
    setAttachmentToDelete(null);
  };

  const getFileIcon = (mimeType: string): string => {
  if (mimeType.startsWith('image/')) return 'Photo2';          // iconName: "Photo2"
  if (mimeType.includes('pdf')) return 'PDF';                  // iconName: "PDF"
  if (mimeType.includes('word')) return 'WordDocument';        // iconName: "WordDocument"
  if (mimeType.includes('excel')) return 'ExcelDocument';      // iconName: "ExcelDocument"
  if (mimeType.includes('powerpoint')) return 'PowerPointDocument'; // iconName: "PowerPointDocument"
  if (mimeType.startsWith('video/')) return 'Video';           // iconName: "Video"
  return 'Page';                                               // fallback icon
};


  const formatFileSize = (bytes: number): string => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  const renderPreviewContent = () => {
    if (!selectedAttachment) return null;

    if (selectedAttachment.mimeType.startsWith('image/')) {
      return (
        <Image
          src={`data:${selectedAttachment.mimeType};base64,${selectedAttachment.documentBody}`}
          alt={selectedAttachment.fileName}
          width="100%"
        />
      );
    }

    if (selectedAttachment.mimeType.includes('pdf')) {
      return (
        <iframe
          src={`data:${selectedAttachment.mimeType};base64,${selectedAttachment.documentBody}`}
          width="100%"
          height="600px"
          title={selectedAttachment.fileName}
        />
      );
    }

    return <Text>Preview not available. Download to view.</Text>;
  };

  useEffect(() => {
    if (message) {
      const timer = setTimeout(() => setMessage(null), 5000);
      return () => clearTimeout(timer);
    }
  }, [message]);

  return (
    <div
      className="attachment-preview-container"
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      style={{
        border: dragOver ? '2px dashed #0078d4' : '1px solid #edebe9',
        borderRadius: '4px',
        padding: '16px',
        minHeight: '400px'
      }}
    >
      {message && (
        <MessageBar
          messageBarType={message.type}
          isMultiline={false}
          onDismiss={() => setMessage(null)}
        >
          {message.text}
        </MessageBar>
      )}

      <CommandBar items={commandBarItems} />

      {isLoading ? (
        <Spinner size={SpinnerSize.large} label="Loading attachments..." />
      ) : (
        <>
          {attachments.length === 0 ? (
            <Stack horizontalAlign="center" verticalAlign="center" style={{ minHeight: '200px' }}>
              <Text variant="large">No attachments found</Text>
              <Text>Drag and drop files or click Upload Files</Text>
            </Stack>
          ) : (
            <DetailsList
              items={attachments}
              columns={columns}
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={0}
              isHeaderVisible={true}
              getKey={(item: IAttachment, index?: number) => item.annotationId}
            />
          )}
        </>
      )}

      <Panel
        isOpen={showPanel}
        onDismiss={() => setShowPanel(false)}
        type={PanelType.large}
        headerText={selectedAttachment?.fileName ?? 'Preview'}
        onRenderFooterContent={() => (
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              text="Download"
              onClick={() => {
                if (selectedAttachment) {
                  void onPreview(selectedAttachment);
                }
              }}
            />
            <DefaultButton text="Close" onClick={() => setShowPanel(false)} />
          </Stack>
        )}
      >
        {renderPreviewContent()}
      </Panel>

      <Dialog
        hidden={!showDeleteDialog}
        onDismiss={() => setShowDeleteDialog(false)}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Delete Attachment',
          subText: `Are you sure you want to delete "${attachmentToDelete?.fileName}"?`
        }}
      >
        <DialogFooter>
          <PrimaryButton text="Delete" onClick={() => void handleDeleteConfirm()} />
          <DefaultButton text="Cancel" onClick={() => setShowDeleteDialog(false)} />
        </DialogFooter>
      </Dialog>
    </div>
  );
};
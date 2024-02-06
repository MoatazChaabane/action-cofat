import * as React from 'react';
import * as Modal from 'react-modal';
import styles from './Actions.module.scss';
import type { IActionsProps } from './IActionsProps';
import { SPHttpClient,SPHttpClientResponse } from '@microsoft/sp-http';
import './modalStyles.scss';

interface IExtractParameters {
  actualiteId: string | null;
  cofatList: string | null;
}

interface IActionsState {
  actualiteId: string | null;
  cofatList: string | null;
  isModalOpen: boolean;
  editedData: {
    title: string;
    content: string;
    attachment: File | null;
  };
}

class Actions extends React.Component<IActionsProps, IActionsState> {
  constructor(props: IActionsProps) {
    super(props);
    this.state = {
      actualiteId: null,
      cofatList: null,
      isModalOpen: false,
      editedData: {
        title: '',
        content: '',
        attachment: null,
      },
    };
  }

  componentDidMount() {
    const { actualiteId, cofatList } = this.extractParametersFromCurrentUrl();
    this.setState({
      actualiteId,
      cofatList,
    });
  }

  private extractParametersFromCurrentUrl(): IExtractParameters {
    const urlSearchParams = new URLSearchParams(window.location.search);
    const actualiteId = this.extractIdFromParameter(urlSearchParams.get('ActualiteId'));
    const cofatList = this.extractListFromParameter(urlSearchParams.get('CofatList'));

    return {
      actualiteId: actualiteId,
      cofatList: cofatList,
    };
  }

  private extractIdFromParameter(parameter: string | null): string | null {
    const match = parameter ? parameter.match(/\b(\d+)\b/) : null;
    return match ? match[1] : null;
  }

  private extractListFromParameter(parameter: string | null): string | null {
    return parameter || null;
  }

  private openModal = () => {
    this.setState({ isModalOpen: true });
  };

  private closeModal = () => {
    this.setState({ isModalOpen: false });
  };

  private updateEditedData = (field: string, value: string | File | null) => {
    this.setState((prevState) => ({
      editedData: {
        ...prevState.editedData,
        [field]: value,
      },
    }));
  };

  
  
  private deleteAllAttachments(listName: string, itemId: number): Promise<void> {
    const { context } = this.props;
  
    const deleteAttachmentsEndpoint = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemId})/AttachmentFiles`;
    console.log(deleteAttachmentsEndpoint);
  
    return context.spHttpClient.get(deleteAttachmentsEndpoint, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        } else {
          console.log(`Error getting attachments: ${response.statusText}`);
          throw new Error(`Error getting attachments: ${response.statusText}`);
        }
      })
      .then((attachments: any) => {
        if (attachments && attachments.value && attachments.value.length > 0) {
          const deletePromises: Promise<void>[] = [];
  
          attachments.value.forEach((attachment: any) => {
            const fileName = attachment.FileName;
            const etag = attachment['odata.etag']; // Fetching the ETag
  
            const deleteAttachmentEndpoint = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items(${itemId})/AttachmentFiles('${fileName}')`;
  
            deletePromises.push(
              context.spHttpClient.post(deleteAttachmentEndpoint, SPHttpClient.configurations.v1, {
                headers: {
                  'X-HTTP-Method': 'DELETE',
                  'IF-MATCH': etag, // Providing the ETag for the request
                },
              })
                .then((deleteResponse: SPHttpClientResponse) => {
                  if (!deleteResponse.ok) {
                    if (deleteResponse.status !== 409) {
                      console.log(`Error deleting attachment: ${deleteResponse.statusText}`);
                      throw new Error(`Error deleting attachment: ${deleteResponse.statusText}`);
                    } else {
                      console.log(`Conflict deleting attachment: ${fileName}`);
                    }
                  }
                })
            );
          });
  
          return Promise.all(deletePromises);
        }
      });
  }
  
  
  
  

  private editItem = async (): Promise<void> => {
    try {
      const { context } = this.props;
      const { actualiteId, cofatList, editedData } = this.state;

      if (!context || !actualiteId || !cofatList) {
        console.error('Context, actualiteId, or cofatList is not defined.');
        return;
      }

      // Construct the URL for the specific item
      const itemUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${cofatList}')/items(${actualiteId})`;

      // Create a plain JavaScript object to represent the item data
      const itemData: {
        Title: string;
        Contenu: string;
        CouleurTexte: string;
      } = {
        Title: editedData.title,
        Contenu: editedData.content,
        CouleurTexte: 'aa', // Remplacez par la valeur appropriée
      };

      // Get the request digest value
      const digestResponse = await context.spHttpClient.post(
        `${context.pageContext.web.absoluteUrl}/_api/contextinfo`,
        SPHttpClient.configurations.v1
      );

      const digest = await digestResponse.json();

      // Headers for the request to update item data
      const itemHeaders = {
        'Accept': 'application/json',
        'Content-Type': 'application/json',
        'X-HTTP-Method': 'MERGE', // Use MERGE to update an existing item
        'X-RequestDigest': digest.FormDigestValue,
        'IF-MATCH': '*', // Include the IF-MATCH header if required
      };

      // Make the request to update item data using fetch
      const updateItemResponse = await context.spHttpClient.post(itemUrl, SPHttpClient.configurations.v1, {
        headers: itemHeaders,
        body: JSON.stringify(itemData),
      });

      // Check if the update was successful
      if (!updateItemResponse.ok) {
        const errorText = await updateItemResponse.text();
        throw new Error(`Update Item Error: ${errorText}`);
      }

      // If an attachment exists, delete existing attachments and upload the new one
      if (editedData.attachment) {
        // Get the existing attachments
      

        // Check if 'results' property exists before iterating
        
          console.log("famma tsawer");
          this.deleteAllAttachments(cofatList, parseInt(actualiteId, 10));
      

        // Upload the new attachment
        const uploadAttachmentUrl = `${itemUrl}/AttachmentFiles/add(FileName='${editedData.attachment.name}')`;
        const uploadAttachmentResponse = await context.spHttpClient.post(
          uploadAttachmentUrl,
          SPHttpClient.configurations.v1,
          {
            headers: {
              'Accept': 'application/json',
              'X-RequestDigest': digest.FormDigestValue,
            },
            body: editedData.attachment,
          }
        );

        // Check if the new attachment upload was successful
        if (!uploadAttachmentResponse.ok) {
          const errorText = await uploadAttachmentResponse.text();
          throw new Error(`Upload Attachment Error: ${errorText}`);
        }
      }

      console.log('Item updated successfully.');

    } catch (error) {
      console.error('Update Item Error:', error);
    }
    this.closeModal();
  };

  private deleteData = async (): Promise<void> => {
    try {
      const { context } = this.props;
      const confirmDelete = window.confirm('Are you sure you want to delete this item?');
  
      if (!context || !confirmDelete) {
        console.error('Context is not defined or delete is canceled.');
        return;
      }
  
      const { actualiteId, cofatList } = this.state;
      const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${cofatList}')/items(${actualiteId})`;
  
      const deleteResponse = await context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
          'Accept': 'application/json',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'DELETE',
        },
      });
  
      if (deleteResponse.ok) {
        console.log('DELETE Response:', deleteResponse);
        
        // Redirect after successful deletion
        window.location.href = context.pageContext.web.absoluteUrl;
        
        // Call the refresh method if the deletion is successful
        // Assuming you have a getData method, uncomment the line below
        // this.getData();
      } else {
        throw new Error(`DELETE Error: ${deleteResponse.statusText}`);
      }
    } catch (error) {
      console.error('DELETE Item Error:', error);
    }
  };
  

  public render(): React.ReactElement<IActionsProps> {
    const { hasTeamsContext } = this.props;
    const { isModalOpen, editedData } = this.state;

    return (
      <section className={`${styles.actions} ${hasTeamsContext ? styles.teams : ''}`}>
        <button onClick={this.openModal}>Edit</button>
        <button onClick={this.deleteData}>Delete</button>

        <Modal
          isOpen={isModalOpen}
          onRequestClose={this.closeModal}
          contentLabel="Edit Item Modal"
        >
          <label>Title:</label>
          <input
            type="text"
            value={editedData.title}
            onChange={(e) => this.updateEditedData('title', e.target.value)}
          />
          <br />
          <label>Content:</label>
          <textarea
            value={editedData.content}
            onChange={(e) => this.updateEditedData('content', e.target.value)}
          />
          <br />
          <label>Attachment:</label>
          <input
            type="file"
            onChange={(e) => this.updateEditedData('attachment', e.target.files?.[0] || null)}
          />
          <br />
          <button onClick={this.editItem}>Save</button>
          <br />
          <button onClick={this.closeModal}>Cancel</button>
        </Modal>
      </section>
    );
  }
}

export default Actions;

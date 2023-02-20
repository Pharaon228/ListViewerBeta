import * as React from 'react';
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { DetailsList, IColumn, DetailsListLayoutMode, IDetailsRowProps, DetailsRow } from 'office-ui-fabric-react/lib/DetailsList';

import { DefaultButton, IconButton, PrimaryButton } from 'office-ui-fabric-react/lib/Button';

import { IListViewerProps } from './IListViewerProps';
import { getSP } from '../../../pnpjsConfig';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Modal } from 'office-ui-fabric-react';
import modalStyles from './Modal.module.scss';


export interface IListViewerState {
  items: IListItem[];
  isEditMode: boolean;
  selectedItem: IListItem | null;
}


interface IListItem {
  Id: number;
  adraOffice: string;
  adraOfficeUrl:string;
  firstName: string;
  lastName: string;
  jobTitle: string;
  workPhone: string;
  email: string;
  skypeId: string;
  [key: string]: string | number;
}

export default class ListViewer extends React.Component<IListViewerProps, IListViewerState> {
  constructor(props: IListViewerProps) {
    super(props);
    this.state = { items: [], isEditMode: false, selectedItem:null };
  }
  private _columns: IColumn[] = [
    {
      key: 'adraOffice',
      name: 'ADRA Office',
      fieldName: 'adraOffice',
      minWidth: 100,
      maxWidth: 200,
      isResizable: true
    },
    {
      key: 'firstName',
      name: 'First Name',
      fieldName: 'firstName',
      minWidth: 100,
      maxWidth: 200,
      isResizable: true
    },
    {
      key: 'lastName',
      name: 'Last Name',
      fieldName: 'lastName',
      minWidth: 100,
      maxWidth: 200,
      isResizable: true
    },
    {
      key: 'jobTitle',
      name: 'Job Title',
      fieldName: 'jobTitle',
      minWidth: 100,
      maxWidth: 200,
      isResizable: true
    },
    {
      key: 'workPhone',
      name: 'Work Phone',
      fieldName: 'workPhone',
      minWidth: 100,
      maxWidth: 200,
      isResizable: true
    },
    {
      key: 'email',
      name: 'Email',
      fieldName: 'email',
      minWidth: 100,
      maxWidth: 200,
      isResizable: true
    },
    {
      key: 'skypeId',
      name: 'Skype ID',
      fieldName: 'skypeId',
      minWidth: 100,
      maxWidth: 200,
      isResizable: true
    },
    {
      key: 'edit',
      name: '',
      fieldName: '',
      minWidth: 50,
      maxWidth: 50,
      isResizable: false,
      onRender: (item: IListItem) => {
        return (
          <IconButton iconProps={{ iconName: 'Edit' }} onClick={() => this._onEditButtonClick(item)} className={modalStyles.editButton} />
        );
      }
    }

  ];

  public componentDidMount() {
    this._loadListData();
  }

  private async _loadListData() {
    const LIST_NAME = 'ADRA Staff';
    const sp: SPFI = getSP(this.props.context);
    const pageUrl = 'https://adra.sharepoint.com/network/afro';
    const items1 = await sp.web.lists.getByTitle(LIST_NAME).items.select()
    .orderBy('ADRA_x0020_Office', true).orderBy('SortID', false)();
    console.log('ADRA  Items', items1);
    const listItems: IListItem[] = items1
    .map((item: any) => ({
      Id: item.Id,
      adraOffice: item.ADRA_x0020_Office.Description,
      adraOfficeUrl : item.ADRA_x0020_Office.Url,
      firstName: item.FirstName,
      lastName: item.Title,
      jobTitle: item.JobTitle,
      workPhone: item.WorkPhone,
      email: item.Email,
      skypeId: item.SkypeID
    }))
    .filter((item) => item.adraOfficeUrl === pageUrl);

    console.log('ADRA Staff Items', listItems);
    this.setState({ items: listItems });
  }

  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn, ascending: boolean = true): void => {
    const { items } = this.state;
    const newItems = items.slice(0);
    const { key } = column;
  
    newItems.sort((a, b) => {
      if (a[key] < b[key]) {
        return ascending ? -1 : 1;
      }
      if (a[key] > b[key]) {
        return ascending ? 1 : -1;
      }
      return 0;
    });
  
    this.setState({
      items: newItems,
    });
  };

  private _onEditButtonClick = (item: IListItem) => {
    console.log('Editing item:', item);
    this.setState({ selectedItem: item });
    this._showEditDialog();
  }

  private _showEditDialog = () => {
    this.setState({ isEditMode: true });
  }

  private _hideEditDialog = () => {
    this.setState({ isEditMode: false });
  }

  private _onCancelEdit = () => {
    // Reset changes to item and hide dialog
    console.log('Cancelling changes to item:', this.state.selectedItem);
    this.setState({ selectedItem: null });
    this._hideEditDialog();
  };

  private _onSaveEdit: React.MouseEventHandler<HTMLDivElement> = async (event) => {
    const editedItem = this.state.selectedItem;
    console.log('Saving edited item:', editedItem);
    const sp: SPFI = getSP(this.props.context);
    const list = sp.web.lists.getByTitle('ADRA Staff');
    const itemToUpdate = {
      Id: editedItem.Id,
      ADRA_x0020_Office: {
        Description: editedItem.adraOffice,
        Url: editedItem.adraOfficeUrl
      },
      FirstName: editedItem.firstName,
      Title: editedItem.lastName,
      JobTitle: editedItem.jobTitle,
      WorkPhone: editedItem.workPhone,
      Email: editedItem.email,
      SkypeID: editedItem.skypeId
    };
  
    try {
      await list.items.getById(editedItem.Id).update(itemToUpdate);
      await this._loadListData(); // reload data after update
      this._hideEditDialog();
    } catch (error) {
      console.error('Error saving item:', error);
    }
  };

  private _onFieldChange = (field: string, value: any) => {
    const selectedItem = { ...this.state.selectedItem, [field]: value };
    this.setState({ selectedItem });
  };

  private _renderEditForm() {
    const { selectedItem } = this.state;
  
    return (
      <Modal isOpen={true} onDismiss={this._hideEditDialog} isBlocking={false}>
        <div className={modalStyles.modal}>
          <div className={modalStyles.modalHeader}>
            <span>Edit ADRA Staff Member</span>
            <IconButton
              iconProps={{ iconName: "Cancel" }}
              onClick={this._hideEditDialog}
              ariaLabel="Close"
            />
          </div>
          <div className={modalStyles.modalBody}>
            <TextField
              label="ADRA Office"
              value={selectedItem.adraOffice}
              onChange={(event, newValue) => this._onFieldChange('adraOffice', newValue)}
            />
            <TextField
              label="ADRA Office URL"
              value={selectedItem.adraOfficeUrl}
              onChange={(event, newValue) => this._onFieldChange('adraOfficeUrl', newValue)}
            />
            <TextField
              label="First Name"
              value={selectedItem.firstName}
              onChange={(event, newValue) => this._onFieldChange('firstName', newValue)}
            />
            <TextField
              label="Last Name"
              value={selectedItem.lastName}
              onChange={(event, newValue) => this._onFieldChange('lastName', newValue)}
            />
            <TextField
              label="Job Title"
              value={selectedItem.jobTitle}
              onChange={(event, newValue) => this._onFieldChange('jobTitle', newValue)}
            />
            <TextField
              label="Work Phone"
              value={selectedItem.workPhone}
              onChange={(event, newValue) => this._onFieldChange('workPhone', newValue)}
            />
            <TextField
              label="Email"
              value={selectedItem.email}
              onChange={(event, newValue) => this._onFieldChange('email', newValue)}
            />
            <TextField
              label="Skype ID"
              value={selectedItem.skypeId}
              onChange={(event, newValue) => this._onFieldChange('skypeId', newValue)}
            />
          </div>
          <div className={modalStyles.modalFooter}>
            <PrimaryButton text="Save" onClick={this._onSaveEdit} />
            <DefaultButton text="Cancel" onClick={this._onCancelEdit} />
          </div>
        </div>
      </Modal>
    );
  }
  
  private _renderList() {
    console.log('Items', this.state.items);
    return (
        <DetailsList
          items={this.state.items}
          columns={this._columns}
          onColumnHeaderClick={this._onColumnClick}
          layoutMode={DetailsListLayoutMode.justified}
          onRenderRow={(props: IDetailsRowProps | undefined): JSX.Element | null => {
            if (!props) {
              return null;
            }

            const { itemIndex } = props;
            const isEvenRow = itemIndex % 2 === 0;

            return (
              <div
                className={isEvenRow ? modalStyles.evenRow : modalStyles.oddRow}
                role="row"
              >
                <DetailsRow {...props} />
              </div>
            );
          }}
        />
    );
}

  
  public render() {
    return (
      <div>
        {this.state.isEditMode ? this._renderEditForm() : null}
        {this._renderList()}
      </div>
    );
  }
}

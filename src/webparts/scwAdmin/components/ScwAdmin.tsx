import * as React from 'react';
import styles from './ScwAdmin.module.scss';
import { IScwAdminProps } from './IScwAdminProps';
import { IScwAdminState } from './IScwAdminState';
import { IOwner } from './IOwner';
import { escape } from '@microsoft/sp-lodash-subset';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { DetailsList, DetailsListLayoutMode, Selection, IColumn, buildColumns, SelectionMode, ISelection, IDetailsColumnStyles, IDetailsHeaderProps, DetailsHeader, DetailsRow, IDetailsRowStyles, } from '@fluentui/react';

import { IContextualMenuProps, IContextualMenuItem, ITooltipHostProps, mergeStyleSets, ScrollbarVisibility, ScrollablePane, StickyPositionType, } from '@fluentui/react';
import ItemDetail from './ItemDetail';
import { autobind, ActionButton, IIconProps, FocusZone, FocusZoneDirection, DefaultButton, } from 'office-ui-fabric-react';
import { getTheme } from 'office-ui-fabric-react/lib/Styling';
import { IDetailsColumnRenderTooltipProps, TooltipHost, Sticky, IRenderFunction, IDetailsListProps, ColumnActionsMode, ContextualMenu, DirectionalHint, Checkbox } from '@fluentui/react';
import * as ReactDom from 'react-dom';
import FilterPane from './FilterPane';

export interface IRequestItem {
  key: number;
  title: string;
  reason: string;
  template: string;
  status: string;
  date: string;
  descriptionEn: string;
  descriptionFr: string;
  owners: IOwner[];
  url: string;
  requesterName: string;
  requesterEmail: string;

}
declare global {
  interface Array<T> {
    from(array: Array<T>);
  }
}

const theme = getTheme();
export interface IRequestItemCollection {
  value: IRequestItem[];
}
var requestItems: IRequestItem[] = [];
var owners: IOwner[] = [];
const headerStyle: Partial<IDetailsColumnStyles> = {
  cellTitle: {
    fontSize: 14,
    fontWeight: 600
  }
};




// test lock header
const classNames = mergeStyleSets({
  wrapper: {
    height: '40vh',
    position: 'relative',
    backgroundColor: 'white',
    margin:'10px'
  },
  filter: {
    backgroundColor: 'white',
    paddingBottom: 20,
    maxWidth: 300,
  },
  header: {
    margin: 0,
    backgroundColor: 'white',
  },
  row: {
    display: 'inline-block',
  },
});



export default class ScwAdmin extends React.Component<IScwAdminProps, IScwAdminState> {

  constructor(props: IScwAdminProps) {
    super(props);

    this.onFilterPanel =this.onFilterPanel.bind(this);

    this.state = {
      showDetailScreen: false,
      selectedTitle: '',
      selectedDesEn: '',
      selectedDesFr: '',
      selectedReason: '',
      selectedKey: '',
      itemId: -1,
      owners: [],
      itemStatus: '',
      items: [],
      contextualMenuProps: undefined,
      isPaneOpen: false,
      selectedUrl: '',
      selectedRequesterName: '',
      selectedRequesterEmail: '',
    };
  }

  public columns: IColumn[] = [
    {

      key: 'title',
      name: 'Space Name',
      fieldName: 'title',
      minWidth: 100,
      maxWidth: 400,
      isResizable: true,
    },
    {
      key: 'reason',
      name: 'Reason',
      fieldName: 'reason',
      minWidth: 100,
      maxWidth: 400,
      isResizable: true,
    },
    {
      key: 'template',
      name: 'Template',
      fieldName: 'template',
      minWidth: 100,
    },
    {
      key: 'status',
      name: 'Status',
      fieldName: 'status',

      minWidth: 100,

    },
    {
      key: 'date',
      name: 'Created Date',
      fieldName: 'date',
      minWidth: 100,
    },
  ];

  public onFilterPanel(filterItems: IRequestItem[]){
    console.log("apple filter items ", filterItems);
    this.setState(() => {
      return {
      items: filterItems,
      };
    });
  }

  private onRenderDetailsHeader: IRenderFunction<IDetailsHeaderProps> = (props, defaultRender) => {
    if (!props) {
      return null;
    }
    const onRenderColumnHeaderTooltip: IRenderFunction<IDetailsColumnRenderTooltipProps> = tooltipHostProps => (
      <TooltipHost {...tooltipHostProps} />
    );
    return (
      <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced>
        {defaultRender!({
          ...props,
          onRenderColumnHeaderTooltip,
        })}
      </Sticky>
    );
  }

  public render(): React.ReactElement<IScwAdminProps> {
    return (

      <div className={styles.scwAdmin}>

        <div className={styles.container}>

          {!this.state.showDetailScreen &&
            <div className={styles.row}>
            <span className={styles.title}>SCW Approvals</span>
              <br /><br />
              <span className={styles.description}>Total {this.state.items.length} items.</span>
              <br />
                     
                      
              
            </div>


          }

          {!this.state.showDetailScreen &&

            <div className={classNames.wrapper}>

              <div className={styles.headerClass}>
                <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
                  <DetailsList
                    styles={headerStyle}
                    items={this.state.items}
                    //  onRenderRow={this._onRenderCell}
                    //   selection={this._selection}
                    selectionMode={SelectionMode.none}
                    columns={this.columns}
                    setKey="set"
                    ariaLabelForSelectionColumn="Toggle selection"
                    ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                    checkButtonAriaLabel="Row checkbox"
                    onItemInvoked={this._onItemInvoked}
                    onRenderRow={this._onRenderRow}
                    onRenderDetailsHeader={this.onRenderDetailsHeader}
                 
                  />
                  {this.state.contextualMenuProps && <ContextualMenu {...this.state.contextualMenuProps} />}
                </ScrollablePane>
              </div>
            </div>

          }

          {this.state.showDetailScreen &&
            <ItemDetail
              returnToMainPage={this.showMainScreen}
              context={this.props.context}
              spaceName={this.state.selectedTitle}
              descriptionEn={this.state.selectedDesEn}
              descriptionFR={this.state.selectedDesFr}
              reason={this.state.selectedReason}
              itemId={this.state.itemId}
              owners={this.state.owners}
              itemStatus={this.state.itemStatus}
              url={this.state.selectedUrl}
              requesterName={this.state.selectedRequesterName}
              requesterEmail={this.state.selectedRequesterEmail}
            />
          }


        </div>
      </div>
    );
  }

public unique(arr){
  return arr.filter(function(item, index, arr){
    return arr.indexOf(item, 0) === index;
  });
}

@autobind
public _alertClicked(){
 

  const statusList = this.state.items.map(x => x.status);

   console.log("status list ", this.unique(statusList));
  this.setState({
    items: this.state.items.filter(i => i.status=='Submitted' || i.status=='Rejected'),
    
  });
  console.log('filter items ', this.state.items);
}

  @autobind
  public showMainScreen() {
    
    this._getItems();
    this.setState(() => {
      return {
        ...this.state,
        showDetailScreen: false
      };
    });
  }


  public renderCustomHeaderTooltip(tooltipHostProps: ITooltipHostProps) {
    return (
      <span
        style={{
          fontSize: "20px",
          fontWeight: "bold"
        }}
      >
        {tooltipHostProps.children}
      </span>
    );
  }
// set column style
  private _onRenderRow: IDetailsListProps['onRenderRow'] = props => {
    const customStyles: Partial<IDetailsRowStyles> = {};
    if (props) {
      if (props.itemIndex % 2 != 0) {
        // Every other row renders with a different background color
        customStyles.root = { backgroundColor: theme.palette.themeLighterAlt };
      }

      return <DetailsRow {...props} styles={customStyles} />;
    }
    return null;
  }

  // after select item, set value to state.
  private _onItemInvoked = (item: IRequestItem): void => {
    
    const ownersEmailStr = item.owners.map(x => x.EMail);
    console.log("new array ", ownersEmailStr);
    this.setState({
      showDetailScreen: true,
      selectedTitle: item.title,
      selectedDesEn: item.descriptionEn,
      selectedDesFr: item.descriptionFr,
      selectedReason: item.reason,
      selectedUrl: item.url,
      itemId: item.key,
      owners: ownersEmailStr,
      itemStatus: item.status,
      selectedRequesterName: item.requesterName,
      selectedRequesterEmail: item.requesterEmail
    });
    console.log("key ", this.state.itemId);
    console.log("title ", this.state.selectedTitle);
  }

  //get requests fields
  public async _getItems(): Promise<void> {
    let items=    await sp.web.lists.getByTitle('Space Requests').items.select('ID', 'Space_x0020_Name', 'Created', 'Template_x0020_Title', 'Team_x0020_Purpose_x0020_and_x00', 'Team_x0020_Alias', "OData__Status", 'Space_x0020_Description_x0020__x', 'Space_x0020_Description_x0020__x0', 'SharePoint_x0020_Site_x0020_URL',"Owner1/EMail","Requester_x0020_email","Requester_x0020_Name").expand("Owner1").getPaged();
    console.log("paged data",items);
    console.log(await (await sp.web.lists.getByTitle('Space Requests').items.getAll()).length);
    while (items.hasNext){
      items = await items.getNext();
      console.log("next page", items);
    }



    
    await sp.web.lists.getByTitle('Space Requests').items.select('ID', 'Space_x0020_Name', 'Created', 'Template_x0020_Title', 'Team_x0020_Purpose_x0020_and_x00', 'Team_x0020_Alias', "OData__Status", 'Space_x0020_Description_x0020__x', 'Space_x0020_Description_x0020__x0', 'SharePoint_x0020_Site_x0020_URL',"Owner1/EMail","Requester_x0020_email","Requester_x0020_Name").expand("Owner1").filter("OData__Status eq 'Submitted' or OData__Status eq 'Approved' or OData__Status eq 'Rejected' or OData__Status eq 'Team Created'").orderBy("Id", true).getAll().then(function (data) {
      console.log("raw data ", data);
      requestItems = [];
      for (var k in data) {
        var longDate: string = data[k].Created;
        var requestItem: IRequestItem = {
          key: data[k].ID,
          title: data[k].Space_x0020_Name,
          reason: data[k].Team_x0020_Purpose_x0020_and_x00,
          template: data[k].Template_x0020_Title,
          status: data[k].OData__Status,
          date: longDate.substring(0, 10),
          descriptionEn: data[k].Space_x0020_Description_x0020__x,
          descriptionFr: data[k].Space_x0020_Description_x0020__x0,
          owners: data[k].Owner1,
          url: data[k].SharePoint_x0020_Site_x0020_URL,
          requesterName: data[k].Requester_x0020_Name,
          requesterEmail: data[k].Requester_x0020_email
        };
        requestItems.push(requestItem);
      }
    
       let item: IRequestItem;

      console.log(requestItems.reverse());
      
    });
    this.setState({
      items: requestItems,
    });
  }

  public componentDidMount() {

    this._getItems();

  }

  private _onColumnClick = (ev: React.MouseEvent<HTMLElement>, column: IColumn): void => {
    if (column.columnActionsMode !== ColumnActionsMode.disabled) {
      this.setState({
        contextualMenuProps: this._getContextualMenuProps(ev, column),
      });
    }
  }

  private onColumnContextMenu = (column: IColumn, ev: React.MouseEvent<HTMLElement>): void => {
    if (column.columnActionsMode !== ColumnActionsMode.disabled) {
      this.setState({
        contextualMenuProps: this._getContextualMenuProps(ev, column),
      });
    }
  }

  private _onContextualMenuDismissed = (): void => {
    this.setState({
      contextualMenuProps: undefined,
    });
  }

  private _getContextualMenuProps(ev: React.MouseEvent<HTMLElement>, column: IColumn): IContextualMenuProps {
    const items = [
      {
        key: 'aToZ',
        name: 'A to Z',
        iconProps: { iconName: 'SortUp' },
        canCheck: true,
        checked: column.isSorted && !column.isSortedDescending,
      },
      {
        key: 'zToA',
        name: 'Z to A',
        iconProps: { iconName: 'SortDown' },
        canCheck: true,
        checked: column.isSorted && column.isSortedDescending,
        onClick: () => this._onSortColumn(column.key, true),
      },
    ];

    return {
      items: items,
      target: ev.currentTarget as HTMLElement,
      directionalHint: DirectionalHint.bottomLeftEdge,
      gapSpace: 10,
      isBeakVisible: true,
      onDismiss: this._onContextualMenuDismissed,
    };
  }

  private _onSortColumn = (columnKey: string, isSortedDescending: boolean): void => {


  }
}



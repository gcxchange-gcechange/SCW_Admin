import * as React from 'react';
import { Panel, PanelType, } from 'office-ui-fabric-react/lib/Panel';
import { Checkbox, Stack, DefaultButton } from '@fluentui/react';
import { autobind } from 'office-ui-fabric-react';
import {IFilterPaneState } from './IFilterPaneState';
import { IFilterPaneProps } from './IFilterPaneProps';

const stackTokens = { childrenGap: 10 };

export default class FilterPane extends React.Component<IFilterPaneProps, IFilterPaneState> {

  constructor(props: IFilterPaneProps) {
    super(props);

    this.state = {
      isOpen: true,
      checkboxs: {
        checkbox1: false,
        checkbox2: false,
        checkbox3: false,
        checkbox4: false,
      },
      items: [],
    };
    this.handle = this.handle.bind(this);
  }

  public render(): React.ReactElement<{}> {
    return (
      <div>
        <Panel
          isOpen={this.state.isOpen}
          type={PanelType.smallFixedFar}
          headerText='Filter by Status'
          closeButtonAriaLabel='Close'
        >
          <Stack tokens={stackTokens}>
            <Checkbox label="Submitted" title="Submitted" onChange={this._onChange} onClick={this.handle}/>
            <Checkbox label="Rejected" title="Rejected" onChange={this._onChange} onClick={this.handle}/>
            <Checkbox label="Approved" title="Approved" onChange={this._onChange} onClick={this.handle}/>
            <Checkbox label="Team Created" title="Team Created" onChange={this._onChange} onClick={this.handle}/>
          </Stack>
          <br />
          <DefaultButton text="Apply" onClick={this._alertClicked} allowDisabledFocus />
        </Panel>
      </div>
    );
  }

  private handle(event) {
    
    let value = event.target.value;
    let { checkboxs } = this.state;
    checkboxs[value] = event.target.checked;   
    this.setState({ checkboxs });
    console.log("value ", this.state.checkboxs);
  }

  private _onChange(ev: React.FormEvent<HTMLElement>, isChecked: boolean) {
    console.log(`The selection ${ev.currentTarget.title} is ${isChecked}`);
  }



  @autobind
  public _alertClicked() {
    console.log("filter item is ",this.props.allItems);
    let filterItems = this.props.allItems.filter(i => i.status=='Submitted' || i.status=='Rejected');
    this.setState ({
      isOpen: !this.state.isOpen,
      items: this.props.allItems.filter(i => i.status=='Submitted' || i.status=='Rejected')
    });
    this.props.onApplyFilter(filterItems);
    console.log("filter item is ",filterItems);

  }
}
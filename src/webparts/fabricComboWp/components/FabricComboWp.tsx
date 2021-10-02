import * as React from 'react';
import { IFabricComboWpProps } from './IFabricComboWpProps';
import { ComboBox, IComboBoxOption, IComboBox, PrimaryButton } from 'office-ui-fabric-react/lib/index';
import { Web } from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { getGUID } from "@pnp/common";

var arr = [];
export interface IStates {
  SingleSelect: any;
  MultiSelect: any;
}

export default class FabricUiComboBox extends React.Component<IFabricComboWpProps, IStates> {

  constructor(props) {
    super(props);
    this.state = {
      SingleSelect: "",
      MultiSelect: []
    };
  }

  private async Save() {
    let web = Web(this.props.webURL);
    await web.lists.getByTitle("ComboBoxExample").items.add({
      Title: getGUID(),
      SingleValueComboBox: this.state.SingleSelect,
      MultiValueComboBox: { results: this.state.MultiSelect }

    }).then(i => {
      console.log(i);
    });
    alert("Submitted Successfully");
  }

  public onComboBoxChange = (ev: React.FormEvent<IComboBox>, option?: IComboBoxOption): void => {
    this.setState({ SingleSelect: option.key });
  }

  public onComboBoxMultiChange = async (event: React.FormEvent<IComboBox>, option: IComboBoxOption): Promise<void> => {
    if (option.selected) {
      await arr.push(option.key);
    }
    else {
      await arr.indexOf(option.key) !== -1 && arr.splice(arr.indexOf(option.key), 1);
    }
    await this.setState({ MultiSelect: arr });
  }

  public render(): React.ReactElement<IFabricComboWpProps> {
    return (
      <div>
        <h1>ComboBox Examples</h1>
        <ComboBox
          placeholder="Single Select ComboBox..."
          selectedKey={this.state.SingleSelect}
          label="Single Select ComboBox"
          autoComplete="on"
          options={this.props.singleValueChoices}
          onChange={this.onComboBoxChange}
        />
        <br />
        <ComboBox
          placeholder="Multi Select ComboBox..."
          selectedKey={this.state.MultiSelect}
          label="Multi Select ComboBox"
          autoComplete="on"
          multiSelect
          options={this.props.multiValueChoices}
          onChange={this.onComboBoxMultiChange}
        />
        <div>
          <br />
          <br />
          <PrimaryButton onClick={() => this.Save()}>Submit</PrimaryButton>
        </div>
      </div>
    );
  }
}

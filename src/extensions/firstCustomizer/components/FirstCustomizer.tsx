import * as React from 'react';
import { Log, FormDisplayMode } from '@microsoft/sp-core-library';
import { FormCustomizerContext } from '@microsoft/sp-listview-extensibility';
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IStackTokens, MessageBar, PrimaryButton, Stack, TextField } from 'office-ui-fabric-react';


import styles from './FirstCustomizer.module.scss';

export interface IFirstCustomizerProps {
  context: FormCustomizerContext;
  displayMode: FormDisplayMode;
  onSave: () => void;
  onClose: () => void;
}

//create state
export interface IFirstFormCustomizerState {
  showmessageBar:boolean; //to show/hide message bar on success
  itemObject:any;
 }

 const stackTokens: IStackTokens = { childrenGap: 40 };
const LOG_SOURCE: string = 'FirstCustomizer';

export default class FirstFormCustomizer extends React.Component<IFirstCustomizerProps, IFirstFormCustomizerState> {

  // Example formatting
    

  // constructor to intialize state and pnp sp object.
  constructor(props: IFirstCustomizerProps,state:IFirstFormCustomizerState) {
    super(props);
    this.state = {showmessageBar:false,itemObject:{}};
    sp.setup({
      spfxContext: this.props.context
    });
  }

  public componentDidMount(): void {
    Log.info(LOG_SOURCE, 'React Element: FirstCustomizer mounted');
  }

  public componentWillUnmount(): void {
    Log.info(LOG_SOURCE, 'React Element: FirstCustomizer unmounted');
  }

  private async createNewItem(){
    const iar: any = await sp.web.lists.getByTitle(this.props.context.list.).items.add({
      Title: this.state.itemObject.title + new Date(),
      Description: this.state.itemObject.Desc
    });
    console.log(iar);
    this.setState({showmessageBar:true});
    //this.props.onSave();
  }

  private updateTitleValue(evt) {
    var item = this.state.itemObject;
    item.title = evt.target.value;
    this.setState({
      itemObject: item
    });
  }

  private updateDescriptionValue(evt) {
    var item = this.state.itemObject;
    item.Desc = evt.target.value;
    this.setState({
      itemObject: item
    });
  }

  private async resetControls(){
    var item = this.state.itemObject;
    item.title = "";
    item.Desc = ""
    this.setState({
      itemObject: item
    });
  }

  public render(): React.ReactElement<{}> {
    return <div className={styles.firstCustomizer}> 
       <TextField required onChange={evt => this.updateTitleValue(evt)} value={this.state.itemObject.title} label="Add Title" />
       <TextField required onChange={evt => this.updateDescriptionValue(evt)} value={this.state.itemObject.Desc} label="Add Description" multiline/>

      <br/>

      <Stack horizontal tokens={stackTokens}>
      <PrimaryButton text="Create New Item" onClick={()=>this.createNewItem()}  />
      <PrimaryButton text="Reset" onClick={()=>this.resetControls()}  />
    </Stack>
      
      <br/>
      {this.state.showmessageBar &&
             <MessageBar   onDismiss={()=>this.setState({showmessageBar:false})}
                dismissButtonAriaLabel="Close">
                "Item saved Sucessfully..."
            </MessageBar>
      }
    
    </div>;
  }
}

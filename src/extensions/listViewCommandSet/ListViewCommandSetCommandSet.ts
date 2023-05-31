import { Log } from '@microsoft/sp-core-library'
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility'
import { Dialog } from '@microsoft/sp-dialog'
import { ILang, getLangStrings } from '../../common/helpers/langHelper'
import Constants from '../../common/helpers/const'

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IListViewCommandSetCommandSetProperties {
  // This is an example replace with your own properties
  sampleTextOne: string
  sampleTextTwo: string
}

const LOG_SOURCE: string = 'ListViewCommandSetCommandSet'

export default class ListViewCommandSetCommandSet extends BaseListViewCommandSet<IListViewCommandSetCommandSetProperties> {

  private _locale: string = 'en'
  private _strings: ILang = null
  private _show: boolean = false

  public async onInit(): Promise<void> {
    this._show = this.context.pageContext.list.serverRelativeUrl.split('/').pop() ===  Constants.ListSystemName

    return getLangStrings(this._locale).then((langStrings) => {
      this._strings = langStrings

      const newCommand: Command = this.tryGetCommand('NEW')
      newCommand.visible = this._show
      newCommand.title = this._strings.Extension.ButtonTitles.New
  
      const editCommand: Command = this.tryGetCommand('EDIT')
      editCommand.visible = false
      editCommand.title = this._strings.Extension.ButtonTitles.Edit
  
      const displayCommand: Command = this.tryGetCommand('DISPLAY')
      displayCommand.visible = false
      displayCommand.title = this._strings.Extension.ButtonTitles.Display
  
      this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged)
    }).catch((err) => {console.error(err)})
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'NEW':
        Dialog.alert(`You have pressed: ${this._strings.Extension.ButtonTitles.New}`)
        .catch((err) => {console.error(err)})
        break
      case 'EDIT':
        Dialog.alert(`You have pressed: ${this._strings.Extension.ButtonTitles.Edit}`)
        .catch((err) => {console.error(err)})
        break
      case 'DISPLAY':
        Dialog.alert(`You have pressed: ${this._strings.Extension.ButtonTitles.Display}`)
        .catch((err) => {console.error(err)})
        break
      default:
        throw new Error('Unknown command')
    }
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed')

    const editCommand: Command = this.tryGetCommand('EDIT')
    const displayCommand: Command = this.tryGetCommand('DISPLAY')
    
    if (editCommand && displayCommand) {
      editCommand.visible = this._show && this.context.listView.selectedRows?.length === 1
      displayCommand.visible = this._show && this.context.listView.selectedRows?.length === 1
    }

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange()
  }
}

import { Log } from '@microsoft/sp-core-library'
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility'
import { Dialog } from '@microsoft/sp-dialog'

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
  private _strings: any = null

  public async onInit(): Promise<void> {
    switch (this._locale) {
      case "en":
        this._strings = await import('../../common/lang/en.json')
        break
      default:
        this._strings = await import('../../common/lang/en.json')
        break
    }

    const newCommand: Command = this.tryGetCommand('NEW')
    newCommand.title = this._strings.Extension.ButtonTitles.New

    const editCommand: Command = this.tryGetCommand('EDIT')
    editCommand.visible = false
    editCommand.title = this._strings.Extension.ButtonTitles.Edit

    const displayCommand: Command = this.tryGetCommand('DISPLAY')
    displayCommand.visible = false
    displayCommand.title = this._strings.Extension.ButtonTitles.Display

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged)

    return Promise.resolve()
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'NEW':
        Dialog.alert(`You have pressed: ${this._strings.Extension.ButtonTitles.New}`)
        break
      case 'EDIT':
        Dialog.alert(`You have pressed: ${this._strings.Extension.ButtonTitles.Edit}`)
        break
      case 'DISPLAY':
        Dialog.alert(`You have pressed: ${this._strings.Extension.ButtonTitles.Display}`)
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
      editCommand.visible = this.context.listView.selectedRows?.length === 1
      displayCommand.visible = this.context.listView.selectedRows?.length === 1
    }

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange()
  }
}

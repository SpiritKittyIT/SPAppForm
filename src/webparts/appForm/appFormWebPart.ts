import * as React from 'react'
import * as ReactDom from 'react-dom'
import { Version } from '@microsoft/sp-core-library'
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane'
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base'
import { IReadonlyTheme } from '@microsoft/sp-component-base'

import AppForm, { IAppFormProps } from './components/appForm'

import { SPFI, spfi, SPFx } from "@pnp/sp"
import { LogLevel, PnPLogging } from "@pnp/logging"
import "@pnp/sp/webs"
import "@pnp/sp/lists"
import "@pnp/sp/items"
import "@pnp/sp/batching"

import { GraphFI, graphfi, SPFx as graphSPFx } from "@pnp/graph"
import "@pnp/graph/sites"
import "@pnp/graph/groups"
import "@pnp/graph/members"
import { getLangStrings, ILang } from '../../common/helpers/langHelper'

export interface IAppFormWebPartProps {
  description: string
}

export default class AppFormWebPart extends BaseClientSideWebPart<IAppFormWebPartProps> {

  private _isDarkTheme: boolean = false
  private _environmentMessage: string = ''
  private _locale: string = 'en'
  private _strings: ILang = null
  private _sp: SPFI = null
  private _graph: GraphFI = null

  public render(): void {
    const appForm: React.ReactElement<{}> =
      React.createElement(AppForm, {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        sp: this._sp,
        graph: this._graph
       } as IAppFormProps)

    ReactDom.render(appForm, this.domElement)
  }

  protected async onInit(): Promise<void> {
    this._sp = spfi().using(SPFx(this.context)).using(PnPLogging(LogLevel.Warning))
    this._graph = graphfi().using(graphSPFx(this.context)).using(PnPLogging(LogLevel.Warning))

    getLangStrings(this._locale).then((langStrings) => {
      this._strings = langStrings
    }).catch((err) => {console.error(err)})
    return
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return
    }

    this._isDarkTheme = !!currentTheme.isInverted
    const {
      semanticColors
    } = currentTheme

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null)
      this.domElement.style.setProperty('--link', semanticColors.link || null)
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null)
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement)
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0')
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: this._strings.Webpart.Properties.Description
          },
          groups: [
            {
              groupName: this._strings.Webpart.Properties.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: this._strings.Webpart.Properties.InputFieldLabel
                })
              ]
            }
          ]
        }
      ]
    }
  }
}

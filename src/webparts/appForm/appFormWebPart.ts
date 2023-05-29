import * as React from 'react'
import * as ReactDom from 'react-dom'
import { Version } from '@microsoft/sp-core-library'
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane'
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base'
import { IReadonlyTheme } from '@microsoft/sp-component-base'

import AppForm from './components/appForm'
import { IAppFormProps } from './components/propInterfaces'

export interface IAppFormWebPartProps {
  description: string
}

export default class AppFormWebPart extends BaseClientSideWebPart<IAppFormWebPartProps> {

  private _isDarkTheme: boolean = false
  private _environmentMessage: string = ''
  private _locale: string = 'en'
  private _strings: any = null

  public render(): void {
    const element: React.ReactElement<IAppFormProps> = React.createElement(
      AppForm,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    )

    ReactDom.render(element, this.domElement)
  }

  protected async onInit(): Promise<void> {
    switch (this._locale) {
      case "en":
        this._strings = await import('../../common/lang/en.json')
        break
      default:
        this._strings = await import('../../common/lang/en.json')
        break
    }
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

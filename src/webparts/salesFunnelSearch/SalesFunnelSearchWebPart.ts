import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  PropertyPaneLabel,
  PropertyPaneLink,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { SalesFunnelSearch } from './components/SalesFunnelSearch';
import { ISalesFunnelSearchProps } from './components/ISalesFunnelSearchProps';
import { DisplayMode } from './models/ISalesFunnelItem';

export interface ISalesFunnelSearchWebPartProps {
  /**
   * Controls which result types are visible.
   * 'documents' | 'listItems' | 'both'
   */
  displayMode: DisplayMode;

  /**
   * Whether the search targets a whole site or a specific list/library URL.
   * 'site' | 'url'
   */
  searchScope: 'site' | 'url';

  /**
   * The SharePoint site URL (when searchScope = 'site')
   * OR the list/library URL  (when searchScope = 'url')
   */
  scopeUrl: string;
}

export default class SalesFunnelSearchWebPart extends BaseClientSideWebPart<ISalesFunnelSearchWebPartProps> {

  protected onInit(): Promise<void> {
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<ISalesFunnelSearchProps> = React.createElement(
      SalesFunnelSearch,
      {
        context: this.context,
        displayMode: this.properties.displayMode || 'both',
        searchScope: this.properties.searchScope || 'site',
        scopeUrl: this.properties.scopeUrl || ''
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const displayModeOptions: IPropertyPaneDropdownOption[] = [
      { key: 'both', text: 'Both (Documents & List Items with tabs)' },
      { key: 'documents', text: 'Documents only' },
      { key: 'listItems', text: 'List Items only' }
    ];

    const searchScopeOptions: IPropertyPaneDropdownOption[] = [
      { key: 'site', text: 'Entire Site (all lists & libraries)' },
      { key: 'url', text: 'Specific List or Library URL' }
    ];

    const isSiteScope = !this.properties.searchScope || this.properties.searchScope === 'site';

    return {
      pages: [
        {
          header: {
            description: 'Configure what this web part searches and displays.'
          },
          groups: [
            {
              groupName: '📋 Display Options',
              groupFields: [
                PropertyPaneDropdown('displayMode', {
                  label: 'Content types to display',
                  options: displayModeOptions,
                  selectedKey: this.properties.displayMode || 'both'
                }),
                PropertyPaneLabel('displayModeDescription', {
                  text: this.getDisplayModeDescription()
                })
              ]
            },
            {
              groupName: '🔍 Search Scope',
              groupFields: [
                PropertyPaneDropdown('searchScope', {
                  label: 'Search scope',
                  options: searchScopeOptions,
                  selectedKey: this.properties.searchScope || 'site'
                }),
                PropertyPaneTextField('scopeUrl', {
                  label: isSiteScope
                    ? 'Site URL'
                    : 'List or Library URL',
                  description: isSiteScope
                    ? 'Enter the full URL of the SharePoint site to search (e.g. https://contoso.sharepoint.com/sites/MySite)'
                    : 'Enter the full URL of a specific list or document library (e.g. https://contoso.sharepoint.com/sites/MySite/Shared Documents)',
                  placeholder: isSiteScope
                    ? 'https://contoso.sharepoint.com/sites/MySite'
                    : 'https://contoso.sharepoint.com/sites/MySite/Lists/MyList',
                  multiline: false,
                  resizable: false
                }),
                PropertyPaneLabel('scopeNote', {
                  text: isSiteScope
                    ? '💡 The web part will search all lists and document libraries within this site.'
                    : '💡 The web part will search only the list or library at the URL above.'
                }),
                PropertyPaneLink('learnMore', {
                  href: 'https://learn.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/overview-client-side-web-parts',
                  text: 'Learn more about SPFx web parts',
                  target: '_blank'
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private getDisplayModeDescription(): string {
    switch (this.properties.displayMode) {
      case 'documents':
        return '📄 Only documents will be searched and shown.';
      case 'listItems':
        return '📋 Only list items will be searched and shown.';
      case 'both':
      default:
        return '🗂 Both documents and list items will be searched. Results appear in three tabs: Documents | List Items | All.';
    }
  }
}

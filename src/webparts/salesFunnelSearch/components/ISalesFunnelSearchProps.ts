import { DisplayMode } from '../models/ISalesFunnelItem';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ISalesFunnelSearchProps {
  /**
   * SPFx context for making API calls
   */
  context: WebPartContext;

  /**
   * Controls which content types to search and display.
   * 'documents' = only documents
   * 'listItems' = only list items
   * 'both' = show both with tabs (Documents | List Items | All)
   */
  displayMode: DisplayMode;

  /**
   * The search scope type.
   * 'site'  = search all lists/libraries in the configured site
   * 'url'   = search only the specific list or library at scopeUrl
   */
  searchScope: 'site' | 'url';

  /**
   * The SharePoint site URL (used when searchScope = 'site')
   * OR the list/library URL (used when searchScope = 'url')
   */
  scopeUrl: string;
}

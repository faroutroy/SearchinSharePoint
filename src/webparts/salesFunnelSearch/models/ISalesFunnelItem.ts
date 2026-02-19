export interface ISalesFunnelItem {
  id: string;
  title: string;
  type: 'document' | 'listItem';
  url: string;
  fileType?: string;
  listName?: string;
  siteUrl?: string;
  modifiedDate?: string;
  author?: string;
  summary?: string;
}

export type DisplayMode = 'documents' | 'listItems' | 'both';
export type SearchScope = 'site' | 'url';

export interface ISearchResult {
  documents: ISalesFunnelItem[];
  listItems: ISalesFunnelItem[];
}

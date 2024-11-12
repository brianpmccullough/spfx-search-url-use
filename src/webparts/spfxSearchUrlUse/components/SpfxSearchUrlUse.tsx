import * as React from 'react';

import type { ISpfxSearchUrlUseProps } from './ISpfxSearchUrlUseProps';
import { DetailsList, PrimaryButton, TextField, DetailsListLayoutMode, IColumn, SelectionMode } from '@fluentui/react';

interface ISearchResult {
  title: string;
  path: string;
  site: string;
  summary: string;
}

interface ISpfxSearchUrlUseState {
  url: string;
  results: ISearchResult[];
  error: string;
}

export default class SpfxSearchUrlUse extends React.Component<ISpfxSearchUrlUseProps, ISpfxSearchUrlUseState> {
  private _columns: IColumn[];

  constructor(props: ISpfxSearchUrlUseProps) {
    super(props);
    this._columns = [
      //{ key: 'column1', name: 'Site', fieldName: 'site', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column2', name: 'Title', fieldName: 'title', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column3', name: 'Path', fieldName: 'path', minWidth: 100, maxWidth: 200, isResizable: true },
      { key: 'column4', name: 'Summary', fieldName: 'summary', minWidth: 100, maxWidth: 200, isResizable: true },
    ];

    this.state = {
      url: '',
      results: [],
      error: '',
    };
  }

/**
 * Splits a given string into segments based on specified characters.
 *
 * This method takes a string and splits it into an array of substrings using
 * the characters provided in the `splitOnCharacters` array. It processes
 * the splits iteratively for each character in the array, allowing for
 * multiple levels of splitting.
 *
 * @private
 * @param {string} value - The string to be split into segments.
 * @param {string[]} [splitOnCharacters=[".", "/", "-"]] - An array of characters
 *        to split the string on. Defaults to [".", "/", "-"] if not provided.
 * @returns {string[]} An array of non-empty segments obtained from the split.
 *
 * @example
 * const result = this.splitByAll("example.com/test-string");
 * console.log(result); // ["example", "com", "test", "string"]
 */
  private splitByAll(value: string, splitOnCharacters: string[] = [".", "/", "-"]): string[] {
    let segments: string[] = [];
    splitOnCharacters.forEach((splitOnCharacter) => {
      console.log(splitOnCharacter);
      if (segments.length > 0) {
        const newSegments: string[] = [];
        segments.forEach((segment) => {
          segment.split(splitOnCharacter).filter(Boolean).forEach(s => newSegments.push(s))
        });
        segments = newSegments;
      } else {
        segments = value.split(splitOnCharacter).filter(Boolean);
      }
    });

    segments = segments.filter(segment => !splitOnCharacters.includes(segment))
    console.log(segments);

    return segments;
  }

  private generateOnearQueryForUrl(url: string): string {
    const parsedUrl = new URL(url);

    // Split the hostname (e.g., "abc.sharepoint.com")
    const domainSegments = this.splitByAll(parsedUrl.hostname);

    // Split the pathname (e.g., "/sites/xyz/")
    const pathSegments = this.splitByAll(parsedUrl.pathname)

    // Combine domain and path segments into a single array, further splitting on any dash characters
    const allSegments = [...domainSegments, ...pathSegments];

    // Construct the ONEAR query for all segments
    const buildOnearQuery = (segments: string[]): string => {
      return segments.reduce((acc, segment) => {
        return acc ? `(${acc} ONEAR(n=0) ${segment})` : segment;
      }, '');
    };

    return buildOnearQuery(allSegments);
  }

  private validateUrl = (url: string): boolean => {
    const urlRegex = /^(https?|ftp):\/\/[^\s/$.?#].[^\s]*$/i;
    return urlRegex.test(url);
  }


  private onSearchClick = async (): Promise<void> => {
    const { url } = this.state;
    if (!this.validateUrl(url)) {
      this.setState({ error: 'Please enter a valid URL.' });
      return;
    }

    const onearQuery = this.generateOnearQueryForUrl(url);
    console.log(onearQuery);

    try {
      const response = await this.props.graphClient
        .api('/search/query')
        .version('v1.0')
        .post({
          requests: [
            {
              entityTypes: ['driveItem', 'listItem'],
              query: {
                queryString: `((FileType:aspx) AND (CanvasContent1OWSHTML:${onearQuery}))`
              },
              fields: [
                "title",
                "path",
                "sitePath"
              ],
              from: 0,
              size: 500 // max page size   
            }
          ]
        });

      const searchResults: ISearchResult[] = response?.value[0]?.hitsContainers[0]?.hits?.length > 0 ? response?.value[0]?.hitsContainers[0]?.hits?.map((hit: any) => {
        return {
          title: hit?.resource?.fields?.title,
          path: hit?.resource?.fields?.path,
          site: hit?.resource?.fields?.sitePath,
          summary: hit?.summary
        }
      }) : [];
      this.setState({ results: searchResults })
      console.log('Search Results:', response);
    } catch (error) {
      console.error('Graph search error:', error);
    }
  };


  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  public render() {
    const { url, results, error } = this.state;

    return (
      <div>
        <TextField
          label="Enter URL:"
          placeholder="https://abc.sharepoint.com/sites/whatever"
          onChange={(e: unknown, value: string | undefined) => this.setState({ url: value || '', error: '' })}
          errorMessage={error}
          value={url}
        />
        <PrimaryButton text="Search" onClick={this.onSearchClick} />
        <div>
          <div>Results:</div>

          <DetailsList
            items={results}
            columns={this._columns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            selectionPreservedOnEmptyClick={true}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            checkButtonAriaLabel="select row"
            selectionMode={SelectionMode.none}
          />
        </div>
      </div>
    );
  }
}

import * as React from 'react';
import { IProjectDocumentProps } from './IProjectDocumentProps';
import { IProjectDocumentState } from './IProjectDocumentState';

export default class ProjectDocument extends React.Component<IProjectDocumentProps, IProjectDocumentState> {
  
    public render(): React.ReactElement<IProjectDocumentProps> {
    return (
      <div>
        {/* <CEOProjectTable webPartTitle={this.props.webPartTitle}></CEOProjectTable> */}
      </div>
    );
  }
}

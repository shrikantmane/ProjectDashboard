import * as React from 'react';
import { IProjectTeamConversationProps } from './IProjectTeamConversationProps';
import { IProjectTeamConversationState } from './IProjectTeamConversationState';

export default class ProjectTeamConversation extends React.Component<IProjectTeamConversationProps, IProjectTeamConversationState> {
  
    public render(): React.ReactElement<IProjectTeamConversationProps> {
    return (
      <div>
        {/* <CEOProjectTable webPartTitle={this.props.webPartTitle}></CEOProjectTable> */}
      </div>
    );
  }
}

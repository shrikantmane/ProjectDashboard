import * as React from 'react';
import { IProjectTeamConversationProps } from './IProjectTeamConversationProps';
import { IProjectTeamConversationState } from './IProjectTeamConversationState';

export default class ProjectTeamConversation extends React.Component<IProjectTeamConversationProps, IProjectTeamConversationState> {
  
    public render(): React.ReactElement<IProjectTeamConversationProps> {
    return (
      <div className="col-xs-12 col-sm-4 cardPadding">
        <div className="well recommendedProjects userFeedback">
          <div className="row">
            <div className="col-sm-12 cardHeading">
              <h5>Team Conversation</h5>
            </div>
            <div className="col-sm-12">
              
            </div>
          </div>
        </div>
      </div>
    );   
  }
}

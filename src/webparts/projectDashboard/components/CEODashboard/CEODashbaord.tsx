import * as React from 'react';
import { ICEODashboardProps } from './ICEODashboardProps';
import { ICEODashboardState } from './ICEODashboardState';
import CEOProjectTable from '../CEOProjectsTable/CEOProjectTable';
import CEOProjectTimeLine from '../CEOProjectTimeLine/CEOProjectTimeLine';

export default class CEODashboard extends React.Component<ICEODashboardProps, ICEODashboardState> {
  
    public render(): React.ReactElement<ICEODashboardProps> {
      var tasks = [
      {
            start: '2018-10-01',
            end: '2018-10-08',
            name: 'Redesign website',
            id: "Task 0",
            progress: 20
        },
        {
            start: '2018-10-03',
            end: '2018-10-06',
            name: 'Write new content',
            id: "Task 1",
            progress: 5,
            dependencies: 'Task 0'
        },
        {
            start: '2018-10-04',
            end: '2018-10-08',
            name: 'Apply new styles',
            id: "Task 2",
            progress: 10,
            dependencies: 'Task 1'
        },
        {
            start: '2018-10-08',
            end: '2018-10-09',
            name: 'Review',
            id: "Task 3",
            progress: 5,
            dependencies: 'Task 2'
        },
        {
            start: '2018-10-08',
            end: '2018-10-10',
            name: 'Deploy',
            id: "Task 4",
            progress: 0,
            dependencies: 'Task 2'
        },
        {
            start: '2018-10-11',
            end: '2018-10-11',
            name: 'Go Live!',
            id: "Task 5",
            progress: 0,
            dependencies: 'Task 4',
            custom_class: 'bar-milestone'
        },
        {
          start: '2018-10-12',
          end: '2018-10-13',
          name: 'Go Live!',
          id: "Task 6",
          progress: 0,
          dependencies: 'Task 5',
          custom_class: 'bar-milestone'
      },
      {
        start: '2018-10-13',
        end: '2018-10-15',
        name: 'Go Live!',
        id: "Task 7",
        progress: 0,
        dependencies: 'Task 6',
        custom_class: 'bar-milestone'
    },
    {
      start: '2018-10-11',
      end: '2018-10-11',
      name: 'Go Live!',
      id: "Task 8",
      progress: 0,
      dependencies: 'Task 7',
      custom_class: 'bar-milestone'
  },
  {
    start: '2018-10-11',
    end: '2018-10-11',
    name: 'Go Live!',
    id: "Task 9",
    progress: 0,
    dependencies: 'Task 8',
    custom_class: 'bar-milestone'
  },
  {
    start: '2018-10-11',
    end: '2018-10-11',
    name: 'Go Live!',
    id: "Task 10",
    progress: 0,
    dependencies: 'Task 9',
    custom_class: 'bar-milestone'
  } 
  ];
    return (
      <div>
        {/* <CEOProjectTimeLine tasks= {tasks} ></CEOProjectTimeLine> */}
        <CEOProjectTable></CEOProjectTable>
      </div>
    );
  }
}

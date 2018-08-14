import { Chart } from "./Project";
export interface IProjectPlanState {
  description?: string;
  currentZoom: string;
  chart: Chart;
  statusList: any;
  teamMembers: any;
  scheduleList: string;
  commentList: string;
  showCommentComponent: boolean,
  documentID: number,
  showDocumentComponent: boolean
}
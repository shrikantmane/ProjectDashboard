import { Chart } from "./Project";
export interface IProjectPlanState {
    description?: string;
    currentZoom : string;
    chart : Chart;
    statusList : any;
    teamMembers : any;
  }
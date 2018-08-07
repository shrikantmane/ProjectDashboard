import  { ViewMode , Groups, TimeLineItems }  from "./TimeLine";
import { Items } from "@pnp/sp";

export interface ITimeLineChartProps {
    groups : Array<Groups>,
    items :  Array<TimeLineItems>
  }

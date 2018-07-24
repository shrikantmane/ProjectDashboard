import  { ViewMode , Groups, TimeLineItems }  from "./ProjectTimeLine";
import { Items } from "@pnp/sp";

export interface ICEOProjectTimeLineProps {
    groups : Array<Groups>,
    items :  Array<TimeLineItems>
  }

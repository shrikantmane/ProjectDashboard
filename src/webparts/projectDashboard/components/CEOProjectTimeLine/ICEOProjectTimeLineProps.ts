import Task , { ViewMode }  from "./ProjectTimeLine";

export interface ICEOProjectTimeLineProps {
  tasks: Array<Task>;
  viewMode?: ViewMode;
  onClick?: () => void;
  onDateChange ?: () => void;
  onProgressChange ?: () => void;
  onViewChange?: () => void;
  customPopupHtml?: any
}

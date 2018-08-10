export enum ViewMode {
    QuarterDay = "Quarter Day",
    HalfDay = "Half Day",
    Day = "Day",
    Week = "Week",
    Month = "Month"
}
export class ProjectTimeLine {     
  groups: Array<Groups>;
  items: Array<TimeLineItems>;  
}

export class Groups {     
  id: number;
  title: string;
}

export class TimeLineItems {     
    id: number;
    group: number;
    title: string;
    start_time: Date;
    end_time: Date;
  }
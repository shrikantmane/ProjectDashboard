export enum ViewMode {
    QuarterDay = "Quarter Day",
    HalfDay = "Half Day",
    Day = "Day",
    Week = "Week",
    Month = "Month"
}

export default class ProjectTimeLine {     
    id: number;
    name: string;
    start: string;
    end: string;
    progress?: number;
    custom_class?: string;
    dependencies? :string;      
}
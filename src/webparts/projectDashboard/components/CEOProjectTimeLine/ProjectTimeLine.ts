export enum ViewMode {
    QuarterDay = "Quarter Day",
    HalfDay = "Half Day",
    Day = "Day",
    Week = "Week",
    Month = "Month"
}

export default class Task {     
    id: string;
    name: string;
    start: string;
    end: string;
    progress: number;
    custom_class?: string;
    dependencies? :string;      
}
export interface IProjectViewState{
    project:string;
    startdate:string;
    enddate:string;
    onhold:string;
    owner:any;
    priority:string;
    complexity:string;
    status:string;
    informationlist:any;
    documentlist:any;
    requirementlist:any;
    teammemberlist:any;
    refreshCalender:boolean;
    Risks:string;
    Id:any;
    onholddate:string,
    events:[{
        id?:number,
        title?:string,
        start?:string,
        end?:string,
    }]
    showComponent:boolean;
    calendarList:any;
    imgURL:any;
    statuscolor:any;
    scheduleList:string;
}
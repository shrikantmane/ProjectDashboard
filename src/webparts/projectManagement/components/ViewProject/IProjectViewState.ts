export interface IProjectViewState{
    
    informationlist:any;
    documentlist:any;
    requirementlist:any;
    teammemberlist:any;
    Id:any;
    events:[{
        id?:number,
        title?:string,
        start?:string,
        end?:string,
    }]
    showComponent:boolean;
    calendarList:any;
}
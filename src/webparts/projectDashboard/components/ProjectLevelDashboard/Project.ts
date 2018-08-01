export class Project {
    public Project: string;   
    public DueDate: string;
    public Priority: string;
    public Task_x0020_List: string;
    public Schedule_x0020_List: string;
    public Project_x0020_Document: string;
    public Project_x0020_Team_x0020_Members: string;
    public Project_x0020_Infromation: string;
}

export class Milestone {
    public Title: string;   
    public DueDate: string;
    public Status0: Status;
    public Priority: string;
}

export class Status {
    public ID: string;
    public Status: string;
    public Status_x0020_Color: string;
}

export class Tag {
    public ID: number;
    public Tag: string;
    public Title:string;
    public Color:string;
}

export class File {
    public Name : string;
    public LinkingUri : string;
    public ServerRelativeUrl : string;
}
export class Document {
    public File : File;
    public Owner : Owner;
    public Created : string;
}

export class Owner {
    public ID: number;
    public Title: string;
    public Department : string;
    public EMail: string;
    public PictureURL :string;
    public JobTitle :string;
    public ImgURL : string;
}

export class Team_x0020_Member {
    public ID: number;
    public EMail: string;
    public Title:string;    
}

export class TeamMember {
   public Team_x0020_Member : Owner;
}

export class RoleResponsibility {
    public ID: number;
    public Roles_Responsibility: string;
    public Owner: Owner;
}

export class Task {
    public ID: number;
    public Week : Week;
    public StartDate : string;
    public EndDate : string;
    public Title : string;
    public OwnerName : string;
    public Status : string;
    public AssignedTo : Array<Owner>;
    public Status0: Status;
}


export enum Week {
    CurrentWeek = "Current Week",
    NextWeek = "Next Week",
    Future = "Future",
    Past = "Past",
}

export class Plan {
    public ID: number;    
    public StartDate : string;
    public DueDate : string;
    public Title : string;
    public Duration : string;
    public AssignedTo : Array<Owner>;
    public Status0: Status;
    public Project: Project;
    public ParentID: Parent;
    public PercentComplete : number;
    public Predecessors :Array<Predecessor>;
}

export class Parent {
    public Id :number;
}

export class Predecessor {
    public Id: number;
    public Title : string;
}

export class ChartData {
    public id: number;
    public text : string;
    public start_date : string;
    public attachment : string;
    public status : string;
    public actualDuration : number;
    public duration : number;
    public color? : string;
    public parent? :number;
    public progress :number;
}

export class ChartLink {
    public id: number;
    public source: number;
    public target: number;
    public type: string;
}

export class Chart {
   public data : Array<ChartData>;
   public links : Array<ChartLink>;
   
}
export class AssignedTo {
    public ID: string;
    public Title: string;
    public EMail: string;
    public imgURL: string; 
}

export class Status {
    public ID: string;
    public Status: string;
    public Status_x0020_Color: string;
}

export class CEOProjects {
    public Project_x0020_ID: string;
    public Project: string;
    public Body:string;
    public Priority: string;
    public Status0: Status;
    public StartDate :string;
    public DueDate: string;
    public StatusText: string;
    public OwnerTitle: string;
    public MildStone : MildStones;
    public AssignedTo: Array<AssignedTo>;
    public MildStoneList: Array<MildStones>;
    public TagList: Array<Tags>;
    public TeamMemberList: Array<TeamMembers>;
    public DocumentList: Array<Documents>;
}

export class Project {
    public Title: string;   
    public ID :number;   
}
export class MildStones {
    public Title: string;   
    public StartDate :string;
    public DueDate: string;
    public Status0: Status;
    public Project: Project;
    public AssignedTo: Array<AssignedTo>;
    public Priority: string;
    public Body: string;
}

export class Tags {
    public ID: number;
    public Tags: string;
    public Title:string;
    public Color:string;
}

export class Team_x0020_Member {
    EMail : string;
    ImgUrl: string
}
export class TeamMembers {
    public ID: number;
    public Project: Project;
    public Start_x0020_Date: string;
    public End_x0020_Date: Status;
    public Status :string;
    public Team_x0020_Member: Team_x0020_Member;
}

export class File {
    public Name : string;
    public LinkingUri : string;
}
export class Documents {
    public File : File;
}
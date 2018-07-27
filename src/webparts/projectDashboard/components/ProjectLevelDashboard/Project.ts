export class Project {
    public Project: string;   
    public DueDate: string;
    public Priority: string;
    public Task_x0020_List: string;
    public Schedule_x0020_List: string;
    public Project_x0020_Document: string;
    public Project_x0020_Team_x0020_Members: string;
}

export class Mildstone {
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
}

export class Team_x0020_Member {
    public ID: number;
    public EMail: string;
    public Title:string;    
}

export class TeamMember {
   public Team_x0020_Member : Team_x0020_Member
}

export class RoleResponsibility {
    public ID: number;
    public Roles_Responsibility: string;
    public Owner: Owner;
}

export class Task {
    public ID: number;
}
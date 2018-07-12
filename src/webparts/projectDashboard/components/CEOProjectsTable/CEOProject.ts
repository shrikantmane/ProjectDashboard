export class AssignedTo {
    public ID: string;
    public Title: string;   
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
    public AssignedTo: Array<AssignedTo>;
}
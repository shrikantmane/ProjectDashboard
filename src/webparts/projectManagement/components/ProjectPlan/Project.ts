export class Predecessor {
    public Id: number;
    public Title : string;
}

export class ChartData {
    public id: number;
    public text? : string;
    public start_date? : string;
    public attachment? : string;
    public status? : string;
    public actualDuration? : number;
    public duration? : number;
    public color? : string;
    public parent? :number;
    public progress :number;
    public comments? :string;
    public statusBackgroudColor? : string;
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


export class Plan {
    public ID: number;    
    public StartDate : string;
    public DueDate : string;
    public Title : string;
    public Duration : string;
    public AssignedTo : Array<Owner>;
    public Status0: Status;
   // public Project: Project;
    public ParentID: Parent;
    public PercentComplete : number;
    public Predecessors :Array<Predecessor>;
}
export class Parent {
    public Id :number;
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

export class Status {
    public ID: string;
    public Status: string;
    public Status_x0020_Color: string;
}
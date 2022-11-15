export interface ITimeAccount {
    id: number;
    title: string;
    description: string;
    sapIdentifier: string;
    picture: string;
    base64: string;
    balanceDays: number;
    blanaceHours: number;
    balanceDaysString: string;
    balanceHoursString: string;
  }

export class TimeAccount implements ITimeAccount {
    constructor(
      public id: number,
      public title: string = "",
      public description: string = "",
      public sapIdentifier: string = "",
      public picture: string = "",
      public base64: string= "",
      public balanceDays: number,
      public blanaceHours: number,
    ) { }
  
    public get balanceDaysString(): string {
      return this.balanceDays.toString()
    }

    public get balanceHoursString(): string {
      return this.blanaceHours.toString()
    }

    public get base64OfImage(): string {
      return this.base64
    }
  }
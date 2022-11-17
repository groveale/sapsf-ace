export interface ITimeAccount {
    id: number;
    title: string;
    description: string;
    sapIdentifierTAT: string;
    sapIdentifierTT: string;
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
      public sapIdentifierTAT: string = "",
      public sapIdentifierTT: string = "",
      public picture: string = "",
      public base64: string= "",
      public balanceDays: number,
      public blanaceHours: number,
      public balanceDaysString: string = "",
      public balanceHoursString: string= ""
    ) { }

    public get base64OfImage(): string {
      return this.base64
    }
  }
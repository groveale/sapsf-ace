
export interface ITimeBooked {
    approvalStatus: string;
    endDate: Date;
    startDate: Date;
    timeType: string;
    quantityInHours: number;
    quantityInDays: number;
  }

export interface ITimeBookedResponse {
  timeBookedPast: ITimeBooked[];
  timeBookedUpcoming: ITimeBooked[];
  balanceDays: number
  balanceHours: number
  daysUntilNextLeave: number
}
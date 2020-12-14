export class CustomItem {
  ID: number;
  Title: string;

  OurMarker: string;

  Person_Manager: any[];

  StateId: number;
  StateTitle: string;

  DateCustomItem: Date;
  
  ItemTypeId: number;
  ItemTypeTitle: string;
  
  DateApproval: Date;
  IsApproved: boolean;
  IsRejected: boolean;
  IsForApproval: boolean;
  IsRevoked: boolean;

  RejectionReason: string;
}
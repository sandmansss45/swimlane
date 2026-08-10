import { IProcessStep } from '../models/IProcessStep';
import { IEmployee } from '../models/IEmployee';
import { IRiskStatement } from '../models/IRiskStatement';

export interface IDataService {
  getProcessSteps(): Promise<IProcessStep[]>;
  getEmployees(): Promise<IEmployee[]>;
  getRiskStatements(): Promise<IRiskStatement[]>;
  addProcessStep(step: Omit<IProcessStep, 'id'>): Promise<IProcessStep>;
  updateProcessStep(step: IProcessStep): Promise<void>;
  deleteProcessStep(id: string): Promise<void>;
}

import { IProcessStep } from '../models/IProcessStep';
import { IEmployee } from '../models/IEmployee';
import { IRiskStatement } from '../models/IRiskStatement';

export interface IDataService {
  getProcessSteps(): Promise<IProcessStep[]>;
  getEmployees(): Promise<IEmployee[]>;
  getRiskStatements(): Promise<IRiskStatement[]>;
  addProcessStep(step: Omit<IProcessStep, 'id'>): Promise<IProcessStep>;
  /**
   * Adds several steps in one call, in array order - used by CSV import.
   * Implementations must preserve order (each created row's SharePoint/mock
   * ID assigned after the previous one), since callers rely on the result
   * landing contiguously at the end of the existing dataset for DependsOn
   * row-number math to stay correct.
   */
  addProcessSteps(steps: Array<Omit<IProcessStep, 'id'>>): Promise<IProcessStep[]>;
  updateProcessStep(step: IProcessStep): Promise<void>;
  deleteProcessStep(id: string): Promise<void>;
}

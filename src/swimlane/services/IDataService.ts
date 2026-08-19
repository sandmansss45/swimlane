import { IProcessStep } from '../models/IProcessStep';
import { IEmployee } from '../models/IEmployee';
import { IRiskStatement } from '../models/IRiskStatement';
import { IProcessGroupLabel } from '../models/IProcessGroupLabel';
import { IProgressIdLabel } from '../models/IProgressIdLabel';
import { IProgressIdLock } from '../models/IProgressIdLock';
import { ISwimlaneComment } from '../models/ISwimlaneComment';

export interface IBulkAddStepsResult {
  created: IProcessStep[];
  // `index` is the position in the array passed to addProcessSteps, not a
  // SharePoint/mock ID (the row never got one) - callers that know a
  // CSV-specific concept of row number (see ImportCsvModal) map this back
  // to something a person can actually locate in their source file.
  failed: Array<{ index: number; error: string }>;
}

export interface IDataService {
  getProcessSteps(): Promise<IProcessStep[]>;
  getEmployees(): Promise<IEmployee[]>;
  getRiskStatements(): Promise<IRiskStatement[]>;
  getProcessGroupLabels(): Promise<IProcessGroupLabel[]>;
  addProcessGroupLabel(groupId: string, name: string): Promise<IProcessGroupLabel>;
  // Renaming an existing custom label (one previously created via
  // addProcessGroupLabel) - not for the static, confirmed-real names in
  // apqcHierarchy.ts, which live in code rather than a list.
  updateProcessGroupLabel(id: string, name: string): Promise<void>;
  getProgressIdLabels(): Promise<IProgressIdLabel[]>;
  addProgressIdLabel(progressId: string, name: string): Promise<IProgressIdLabel>;
  updateProgressIdLabel(id: string, name: string): Promise<void>;
  getProgressIdLocks(): Promise<IProgressIdLock[]>;
  // Always creates a NEW record (see IProgressIdLock for why - append-only
  // audit trail, not a mutable status flag).
  lockProgressId(progressId: string, region: string, lockedBy: string, reason: string): Promise<IProgressIdLock>;
  // Sets unlockedBy/unlockedAt on the existing record identified by id -
  // the one and only mutation this list ever gets, and even then only
  // ever fills in previously-blank fields, never overwrites what's there.
  unlockProgressId(id: string, unlockedBy: string, reason: string): Promise<void>;
  getSwimlaneComments(): Promise<ISwimlaneComment[]>;
  // Always creates a NEW record (see ISwimlaneComment - append-only,
  // never edited or deleted once posted).
  addSwimlaneComment(progressId: string, region: string, author: string, comment: string): Promise<ISwimlaneComment>;
  addProcessStep(step: Omit<IProcessStep, 'id'>): Promise<IProcessStep>;
  /**
   * Adds several steps in one call, in array order - used by CSV import.
   * Implementations must preserve order (each created row's SharePoint/mock
   * ID assigned after the previous one), since callers rely on the result
   * landing contiguously at the end of the existing dataset for DependsOn
   * row-number math to stay correct.
   *
   * Never rejects on an individual row failing - a network blip or one bad
   * row used to throw out of the whole batch, silently discarding every
   * row already successfully created (real, permanent items that existed
   * in SharePoint but the caller never found out about) and never even
   * attempting whatever came after it in the file. Every row that
   * succeeds is always returned in `created`; every row that fails is
   * reported in `failed` with its index into the input array and why, so
   * the caller can show the user exactly what did and didn't make it in.
   */
  addProcessSteps(steps: Array<Omit<IProcessStep, 'id'>>): Promise<IBulkAddStepsResult>;
  updateProcessStep(step: IProcessStep): Promise<void>;
  deleteProcessStep(id: string): Promise<void>;
}

import { SYNC_FLAG } from '../types';
export declare function transaction(origin?: SYNC_FLAG): (_target: any, key: string, descriptor: PropertyDescriptor) => PropertyDescriptor;

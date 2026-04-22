import { UserItem } from '../../types';
export type UserInfo = {
    clientId: number;
    users: UserItem[];
    fileId: string;
    fileName: string;
};
type Action = {
    setUsers(users: UserItem[]): void;
    setClientId(clientId: number): void;
    setFileName(name: string): void;
    setFileInfo(id: string, name: string): void;
};
export declare const useUserInfo: import('zustand').UseBoundStore<import('zustand').StoreApi<UserInfo & Action>>;
export {};

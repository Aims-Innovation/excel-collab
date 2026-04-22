import { IWindowSize } from '../types';
export declare const IMAGE_TYPE_MAP: {
    readonly 'image/apng': readonly [".apng"];
    readonly 'image/bmp': readonly [".bmp"];
    readonly 'image/x-icon': readonly [".ico", ".cur"];
    readonly 'image/png': readonly [".png"];
    readonly 'image/webp': readonly [".webp"];
    readonly 'image/svg+xml': readonly [".svg"];
    readonly 'image/avif': readonly [".avif"];
    readonly 'image/gif': readonly [".gif"];
    readonly 'image/jpeg': readonly [".jpeg", ".jpg", ".jfif", ".pjpeg", ".pjp"];
};
export declare function extractImageType(src: string): {
    ext: string;
    type: string;
    base64: string;
};
export declare function convertFileToTextOrBase64(file: File, isBase64?: boolean): Promise<string>;
export declare function getImageSize(src: string): Promise<IWindowSize>;
export declare function convertBase64toBlob(base64: string, contentType?: string): Blob;

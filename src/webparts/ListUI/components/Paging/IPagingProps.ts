
export interface IPagingProps {
    onBackToFirst: () => void;
    onNextPage: (pageNumber: number) => void;
    currentPage: number;
    nextEnabled: boolean;
}

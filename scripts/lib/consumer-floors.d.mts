export interface ConsumerFloors {
  node: string[];
  typescript: string;
  vite: string;
  webpack: string;
  react: string;
  typesReact: string;
}
export const FLOORS: ConsumerFloors;
export function floorProblems(
  stated: {
    coreReadme: string;
    reactReadme: string;
    vueReadme: string;
    engines: string | undefined;
    reactPeer: string | undefined;
    ci: string;
  },
  floors?: ConsumerFloors
): string[];

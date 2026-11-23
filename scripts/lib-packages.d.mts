// Copied from pptx-editor/scripts/lib-packages.d.mts at e22f3f3b (EigenPal). Differences: the
// `private` field and `PUBLISHED_PACKAGES`.
export interface PublishedPackage {
  name: string;
  root: string;
  slug: string;
  entries: string[];
  /** The manifest is private: npm never publishes the package. */
  private?: boolean;
}
export declare const PACKAGES: PublishedPackage[];
export declare const PUBLISHED_PACKAGES: PublishedPackage[];

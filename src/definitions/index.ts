import keywords from "./keywords.json";
import operators from "./operators.json";

import { CompletionItem, CompletionItemKind } from "vscode";

const completions = new Array<CompletionItem>();

for (const entry in keywords) {
  const itm = new CompletionItem(entry, CompletionItemKind.Keyword);
  itm.detail = entry;
  itm.documentation = keywords[entry]?.documentation;
  completions.push(itm);
}

for (const entry in operators) {
  const itm = new CompletionItem(entry, CompletionItemKind.Operator);
  itm.detail = entry;
  itm.documentation = operators[entry]?.documentation;
  itm.filterText = `Operator ${entry}`;
  completions.push(itm);
}

export default completions;

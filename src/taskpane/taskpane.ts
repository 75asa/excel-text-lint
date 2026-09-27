import { readSelectedCells } from "../hosts/excel/selection";

const status = document.getElementById("status") as HTMLParagraphElement;
const button = document.getElementById("read-selection") as HTMLButtonElement;
const list = document.getElementById("cells") as HTMLOListElement;

Office.onReady(({ host }) => {
  if (host !== Office.HostType.Excel) {
    status.textContent = `未対応のホストです: ${host ?? "（Office の外で開かれています）"}`;
    return;
  }
  status.textContent = "セルを選択してボタンを押してください。";
  button.disabled = false;
  button.addEventListener("click", () => void showSelectedCells());
});

async function showSelectedCells(): Promise<void> {
  button.disabled = true;
  try {
    const cells = await readSelectedCells();
    list.replaceChildren(
      ...cells.map(({ address, text }) => {
        const item = document.createElement("li");
        const addr = document.createElement("span");
        addr.className = "address";
        addr.textContent = address;
        const body = document.createElement("span");
        body.className = "text";
        body.textContent = text;
        item.append(addr, body);
        return item;
      }),
    );
    status.textContent = `${cells.length} 件のセルを読み取りました。`;
  } catch (error) {
    console.error(error);
    status.textContent = `読み取りに失敗しました: ${error instanceof Error ? error.message : String(error)}`;
  } finally {
    button.disabled = false;
  }
}

/**
 * Text flow between linked text boxes.
 *
 * Word lets a text box continue in another one: the author draws several boxes,
 * links them into a chain, and types into the first. Text that does not fit
 * carries over to the next box in the chain. In the file only the first box of a
 * chain holds content, so a renderer that draws each box on its own puts the
 * whole text into the first one — where a fixed box height silently cuts it off.
 *
 * The chain is walked over the rendered document and the text is redistributed
 * the way Word lays it out: fill a box, carry the rest over to the next one.
 */

/** Extra pixels a line may reach past a box before it counts as not fitting. */
const TOLERANCE = 1;

/**
 * Redistributes text over the chains of linked text boxes in a rendered document.
 *
 * Measures laid out text, so the document has to be in the page already and the
 * fonts it uses loaded — both are awaited here.
 */
export async function flowLinkedTextboxes(root: ParentNode): Promise<void> {
	const boxes = Array.from(root.querySelectorAll<SVGSVGElement>('svg[data-shape-id]'));
	const chained = boxes.filter(b => b.hasAttribute('data-next-shape'));

	if (chained.length == 0)
		return;

	// a layout pass has to run first: fonts are only requested once something is
	// laid out with them, and `fonts.ready` waits for the requested ones alone —
	// awaited before that, it resolves right away
	chained[0].getBoundingClientRect();

	// splitting against fallback metrics breaks the text at the wrong word, and the
	// result looks like a correct render rather than a failed one
	await (document as any).fonts?.ready;

	const byId = new Map(boxes.map(b => [b.getAttribute('data-shape-id'), b]));
	const continuations = new Set(chained.map(b => b.getAttribute('data-next-shape')));
	const starts = chained.filter(b => !continuations.has(b.getAttribute('data-shape-id')));

	for (const start of starts) {
		flowChain(chainFrom(start, byId));
	}
}

/** Boxes of one chain, from the one holding the text to the last continuation. */
function chainFrom(start: SVGSVGElement, byId: Map<string, SVGSVGElement>): SVGSVGElement[] {
	const chain: SVGSVGElement[] = [];

	for (let box = start; box; box = byId.get(box.getAttribute('data-next-shape'))) {
		if (chain.includes(box)) // a chain closed into a loop would never end
			break;

		chain.push(box);
	}

	return chain;
}

function flowChain(chain: SVGSVGElement[]) {
	const contents = chain.map(contentOf);

	// only the first box of a chain holds text; anything a continuation carries is
	// filler the editor left behind, and Word draws none of it
	for (const content of contents.slice(1)) {
		content?.replaceChildren();
	}

	// heights are taken before anything moves: a box keeps the size it was drawn
	// with, and reading them later would measure boxes the flow has already filled
	const limits = chain.map(box => box.getBoundingClientRect().height);

	for (let i = 0; i < chain.length - 1; i++) {
		if (contents[i] && contents[i + 1])
			carryOver(contents[i], contents[i + 1], limits[i]);
	}

	// whatever does not fit the last box stays visible: cutting it off loses text
	// without saying so
	const last = contents[contents.length - 1];

	if (last)
		last.style.overflow = 'visible';
}

/** The box body: the element the text of a text box is rendered into. */
function contentOf(box: SVGSVGElement): SVGForeignObjectElement | null {
	return box.querySelector('foreignObject');
}

/** Moves the part of `from` that does not fit `limit` to the front of `to`. */
function carryOver(from: SVGForeignObjectElement, to: SVGForeignObjectElement, limit: number) {
	// whether anything overflows is answered by the same measurement that says where
	// to cut: box heights and text positions are only comparable in one space, and
	// under a zoomed sheet `scrollHeight` and `getBoundingClientRect` are not in it
	const point = firstBelow(from, from.getBoundingClientRect().top + limit);

	if (!point)
		return;

	const range = document.createRange();
	range.setStart(point.node, point.offset);
	range.setEnd(from, from.childNodes.length);

	to.insertBefore(range.extractContents(), to.firstChild);

	dropEmpty(from.lastElementChild);
	dropEmpty(to.firstElementChild);
}

/**
 * Where the text leaves the box: the first character whose line does not end
 * above `y`. A line counts as fitting only whole — Word never draws half of one.
 *
 * Characters go down the box in order, so the crossing is found by halving the
 * text rather than measuring every character.
 */
function firstBelow(container: Element, y: number): { node: Text, offset: number } | null {
	const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT);
	const nodes: { node: Text, start: number }[] = [];
	let length = 0;

	for (let node: Node; (node = walker.nextNode());) {
		nodes.push({ node: node as Text, start: length });
		length += node.nodeValue.length;
	}

	if (length == 0)
		return null;

	const range = document.createRange();

	const bottomOf = (index: number) => {
		const at = locate(nodes, index);
		range.setStart(at.node, at.offset);
		range.setEnd(at.node, at.offset + 1);

		const box = range.getBoundingClientRect();

		// a character a line break fell on has no box of its own
		return box.height == 0 ? null : box.bottom;
	};

	let low = 0, high = length - 1, found: number = null;

	while (low <= high) {
		const middle = (low + high) >> 1;
		const bottom = bottomOf(middle);

		if (bottom == null || bottom <= y + TOLERANCE) {
			low = middle + 1;
		} else {
			found = middle;
			high = middle - 1;
		}
	}

	return found == null ? null : locate(nodes, found);
}

/** Text node and offset holding the character at `index` of the whole text. */
function locate(nodes: { node: Text, start: number }[], index: number): { node: Text, offset: number } {
	let found = nodes[0];

	for (const entry of nodes) {
		if (entry.start > index)
			break;

		found = entry;
	}

	return { node: found.node, offset: index - found.start };
}

/**
 * Removes the empty paragraph a split leaves behind.
 *
 * Splitting inside a paragraph clones it on both sides, so the box the text left
 * ends with an empty paragraph and the box it moved to starts with one. Left in
 * place, each takes a line of a box that is one line tall. Only those two ends
 * are looked at: an empty paragraph elsewhere is one the author typed.
 */
function dropEmpty(element: Element | null) {
	if (element?.textContent.trim() == '' && !element?.querySelector('img, svg'))
		element?.remove();
}

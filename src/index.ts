import xlsx from "node-xlsx";
import path from "node:path";
import fs from "node:fs";

const fileName = "kereto-first.xlsx";
const [{ data }] = xlsx.parse(path.join(__dirname, "../xlsx-files/" + fileName));

const dataSorted = data.sort((a, b) => b[2] + b[3] - (a[2] + a[3]));

const dataAgg = dataSorted.map(([id, name, grades1 = 0, grades2 = 0], index) => {
	return {
		id,
		name,
		total: grades1 + grades2,
		grades1: { rank: dataSorted.filter((x) => x[2] > grades1).length+1, value: grades1 },
		grades2: { rank: dataSorted.filter((x) => x[3] > grades2).length+1, value: grades2 },
		totalRank: dataSorted.filter((x) => x[2] + x[3] > grades1 + grades2).length+1,
	};
});

console.table(dataAgg);

const resultData = [
	[
		"م",
		"الاسم",
		"مجموع الدرجات",
		"درجة فيزياء الحرارية",
		"درجة فيزياء الخواص",
		"ترتيب فيزياء الحرارية",
		"ترتيب فيزياء الخواص",
		"الترتيب على الدفعة",
	],
	...dataAgg
		.map((row) => [
			row.id,
			row.name,
			row.total,
			row.grades1.value,
			row.grades2.value,
			row.grades1.rank,
			row.grades2.rank,
			row.totalRank,
		])
		.filter((x) => x[0])
		.sort((a, b) => a[0] - b[0]),
];

const resultBuffer = xlsx.build([{ name: "ترتيب الدفعة (حبق اكريتو)", data: resultData, options: { } }], { }); // Returns a buffer
fs.writeFileSync(path.join(__dirname, `../out/tarteeb-jadwal.xlsx`), resultBuffer);

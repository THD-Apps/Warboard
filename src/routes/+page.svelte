<script>
	import ExcelJS from 'exceljs';
	import moment from 'moment';

	let finishedData = null;
	let loading = false;

	let reportDate = '';
	let reportUser = '';

	function startLoad(event) {
		loading = true;
		setTimeout(() => {
			processFile(event);
		}, 1000);
	}

	function removeDuplicateRows(jsonData) {
		const uniqueRows = new Set();

		// Filter out duplicates
		const filteredData = jsonData.filter((row) => {
			const rowString = JSON.stringify(row);
			if (uniqueRows.has(rowString)) {
				// Duplicate, return false to filter it out
				return false;
			} else {
				// Unique, add to the set and return true to keep it
				uniqueRows.add(rowString);
				return true;
			}
		});

		return filteredData;
	}

	function reorganizeData(jsonData) {
		const chunkRanges = [
			{ start: 0, end: 49 },
			{ start: 50, end: 95 },
			{ start: 96, end: 146 },
			{ start: 147, end: 197 }
		];

		const reorderedData = [];

		for (const range of chunkRanges) {
			const chunkData = jsonData.slice(range.start, range.end + 1);
			const chunkRows1 = [];
			const chunkRows2 = [];

			for (const row of chunkData) {
				// First 10 columns
				const chunkRow1 = [];
				for (let col = 0; col < 10; col++) {
					chunkRow1.push(row[col]);
				}
				chunkRows1.push(chunkRow1);

				// Columns 12-18
				const chunkRow2 = [];
				for (let col = 11; col < 18; col++) {
					chunkRow2.push(row[col]);
				}
				chunkRows2.push(chunkRow2);
			}

			reorderedData.push(chunkRows1, chunkRows2);
		}

		// Flatten the array of arrays
		return reorderedData.flat();
	}

	function findColumn6Value(jsonData) {
		const foundRow = jsonData.find((row) => row[3] === 'Printed for :');

		if (foundRow) {
			const column6Value = foundRow[5];
			return column6Value;
		} else {
			return 'Unknown';
		}
	}

	function updateColumn0ForPattern(jsonData) {
		const updatedData = jsonData.map((row) => {
			// Check if column 2 matches the specified pattern
			if (/^\d{3} -/.test(row[1])) {
				// If yes, update column 0 to "DEPT"
				row[0] = 'DEPT';
			}
			return row;
		});

		return updatedData;
	}

	function updateColumn0ForEquality(jsonData) {
		const updatedData = jsonData.map((row) => {
			// Check if columns 2, 3, 4, and 5 are the same and column 0 is not "DEPT"
			if (row[0] !== 'DEPT' && row[1] === row[2] && row[2] === row[3] && row[3] === row[4]) {
				// If yes, update column 0 to "JOB"
				row[0] = 'JOB';
			}
			return row;
		});

		return updatedData;
	}

	function unnestRichText(jsonData) {
		const unnestedData = jsonData.map((row) => row.map((cell) => unnestCell(cell)));
		return unnestedData;
	}

	function unnestCell(cell) {
		if (cell && typeof cell === 'object' && 'richText' in cell) {
			// If the cell is an object with 'richText' property, unnest it
			return unnestRichTextValue(cell.richText);
		}
		return cell;
	}

	function unnestRichTextValue(richText) {
		return richText.map((segment) => segment.text).join('');
	}

	function changeFirstColumnShifts(jsonData) {
		const modifiedData = jsonData.map((row) => {
			// Check if the row has 7 columns and column 3 contains a time range pattern
			if (row.length === 10 && /\d{1,2}:\d{2}(AM|PM)-\d{1,2}:\d{2}(AM|PM)/.test(row[4])) {
				// Update column 0 to "SHIFT"
				row[0] = 'SHIFT';
				row[1] = row[1];
				// Update column 1 to the value of column 5
				row[2] = row[4];
				row[3] = row[7];
				// Keep only columns 0, 1, and 2
				return row.slice(0, 4);
			}
			return row;
		});

		return modifiedData;
	}

	function changeSecondColumnShifts(jsonData) {
		const modifiedData = jsonData.map((row) => {
			// Check if the row has 7 columns and column 3 contains a time range pattern
			if (row.length === 7 && /\d{1,2}:\d{2}(AM|PM)-\d{1,2}:\d{2}(AM|PM)/.test(row[2])) {
				// Update column 0 to "SHIFT"
				row[0] = 'SHIFT';
				row[1] = row[1];
				// Update column 1 to the value of column 5
				row[2] = row[2];
				row[3] = row[4];
				// Keep only columns 0, 1, and 2
				return row.slice(0, 4);
			}
			return row;
		});

		return modifiedData;
	}

	function filterRows(jsonData) {
		const filteredData = jsonData.filter((row) => {
			const firstColumnValue = row[0];
			return (
				firstColumnValue === 'SHIFT' || firstColumnValue === 'DEPT' || firstColumnValue === 'JOB'
			);
		});

		return filteredData;
	}

	function nestData(jsonData) {
		let nestedData = [];
		let currentDept = null;
		let currentJob = null;

		for (const row of jsonData) {
			const type = row[0];

			if (type === 'DEPT') {
				// New department, reset currentJob
				currentDept = { name: row[1], jobs: [] };
				currentJob = null;
				nestedData.push({ ...row, ...currentDept });
			} else if (type === 'JOB') {
				// New job, reset currentJob
				currentJob = { name: row[1], shifts: [] };
				currentDept.jobs.push({ ...row, ...currentJob });
			} else if (type === 'SHIFT') {
				// New shift, add to currentJob or create a default "Associate" job if there is no currentJob
				if (currentJob) {
					currentJob.shifts.push({ time: row[2], ...row });
				} else {
					currentJob = { name: 'Associate', shifts: [{ time: row[2], ...row }] };
					currentDept.jobs.push(currentJob);
				}
			}
		}

		// Remove jobs with no children
		nestedData = nestedData.filter(
			(item) => item[0] === 'DEPT' || item[0] === 'JOB' || item[0] === 'SHIFT'
		);

		return nestedData;
	}

	function convertToMilitaryTime(timeString) {
		// Parse the input time string
		const parsedTime = /(\d{1,2}):(\d{2})([APMapm]{2})/.exec(timeString);

		if (parsedTime) {
			let hours = parseInt(parsedTime[1], 10);
			const minutes = parsedTime[2];
			const period = parsedTime[3].toUpperCase();

			// Adjust hours based on AM/PM
			if (period === 'PM' && hours !== 12) {
				hours += 12;
			} else if (period === 'AM' && hours === 12) {
				hours = 0;
			}

			// Format hours and minutes as two digits
			const formattedHours = hours.toString().padStart(2, '0');
			return `${formattedHours}:${minutes}`;
		}

		// Return the input string if it doesn't match the expected format
		return timeString;
	}

	function processFile(event) {
		const file = event.target.files[0];

		if (file) {
			const reader = new FileReader();

			reader.onload = async function (event) {
				const arrayBuffer = event.target.result;

				// Use exceljs to read the Excel file
				const workbook = new ExcelJS.Workbook();
				await workbook.xlsx.load(arrayBuffer);

				const worksheet = workbook.worksheets[0];
				let jsonData = [];

				worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
					const rowJson = {};
					row.eachCell((cell, colNumber) => {
						rowJson[colNumber] = cell.value;
					});
					jsonData.push(rowJson);
				});
				jsonData = reorganizeData(jsonData);
				jsonData = removeDuplicateRows(jsonData);
				reportDate = jsonData[1][3].substring(0, 10);
				reportUser = findColumn6Value(jsonData);
				jsonData = updateColumn0ForPattern(jsonData);
				jsonData = updateColumn0ForEquality(jsonData);
				jsonData = unnestRichText(jsonData);
				jsonData = changeSecondColumnShifts(jsonData);
				jsonData = changeFirstColumnShifts(jsonData);
				jsonData = filterRows(jsonData);
				jsonData = jsonData.slice(1);
				jsonData = nestData(jsonData);
				const foundItem = jsonData.find((item) => item[1] === '096 - Lot');

				// If the item is found, filter the jobs property
				if (foundItem) {
					foundItem.jobs = foundItem.jobs.filter((job) => {
						const property1Value = job[1] ?? job.name;
						return property1Value === 'Associate' || property1Value === 'Facility';
					});
				}
				console.log(jsonData);
				finishedData = jsonData;
				loading = false;
			};

			reader.readAsArrayBuffer(file);
		}
	}
</script>

<div class="text-center">
	{#if loading}
		<div class="font-medium mt-5 text-[25px]">
			Hang tight! Our data hamsters are running as fast as they can...
		</div>
		<div class="hamster text-[200px]">🐹</div>
	{:else if finishedData}
		<div class="flex justify-center screen-only">
			<div class="w-2/5 grid grid-cols-2 gap-3">
				<div
					class="cursor-pointer text-white my-3 bg-blue-700 hover:bg-blue-800 focus:outline-none focus:ring-4 focus:ring-blue-300 font-medium rounded-full text-sm w-50 py-1.5 text-center mb-2 dark:bg-blue-600 dark:hover:bg-blue-700 dark:focus:ring-blue-800"
					on:click={() => window.print()}
				>
					Print Report
				</div>
				<div
					class="cursor-pointer text-white my-3 bg-blue-700 hover:bg-blue-800 focus:outline-none focus:ring-4 focus:ring-blue-300 font-medium rounded-full text-sm w-50 py-1.5 text-center mb-2 dark:bg-blue-600 dark:hover:bg-blue-700 dark:focus:ring-blue-800"
					on:click={() => window.location.reload()}
				>
					Start Over
				</div>
			</div>
		</div>
		<div class="header mx-[20%] print:mx-0">
			<div class="grid grid-cols-3 text-[14px] mb-3">
				<div class="text-left">Date: {moment(reportDate).format('dddd, MMMM Do YYYY')}</div>
				<div class="font-bold">Daily Break Schedule</div>
				<div class="text-right">Generated By: {reportUser}</div>
			</div>
		</div>
		<div
			class="font-bold grid grid-cols-12 border-b border-t border-black text-[20px] print:text-[14px] mx-[20%] print:mx-0"
		>
			<div class="col-span-5 grid grid-cols-12">
				<div class="px-2 border-l border-black uppercase col-span-6 py-1">Associate</div>
				<div class="px-2 border-l border-black uppercase col-span-6 py-1">Shift</div>
			</div>
			<div class="col-span-7 grid grid-cols-12">
				<div class="col-span-7 grid grid-cols-3">
					<div class="px-2 border-l border-black uppercase py-1">Break</div>
					<div class="px-2 border-l border-black uppercase py-1">Lunch</div>
					<div class="px-2 border-l border-black uppercase py-1">Break</div>
				</div>
				<div class="col-span-5">
					<div class="px-2 border-l border-black border-r uppercase py-1">Task</div>
				</div>
			</div>
		</div>
		<div class="text-[18px] print:text-[14px] mx-[20%] print:mx-0" id="display-area">
			<div class="text-left">
				{#each finishedData as department}
					<div class="column-break" class:mb-[15px]={department.jobs.length}>
						{#if department.jobs.length}
							<div class="border border-black bg-gray-800 text-white text-center font-bold">
								<!-- DEPARTMENT NAME -->
								{department[1]}
							</div>
						{/if}

						{#each department.jobs as job}
							{#if job.shifts.length}
								{#if department.jobs.length > 1}
									<!-- JOB TITLE -->
									<div
										class="border-l border-r border-b border-black bg-gray-400 text-center font-bold"
									>
										{job.name ?? job[1]}
									</div>
								{/if}
								{#each job.shifts as shift}
									{@const diff = moment(
										'2023-12-12 ' + convertToMilitaryTime(shift[2].split('-')[1])
									).diff(
										moment('2023-12-12 ' + convertToMilitaryTime(shift[2].split('-')[0])),
										'minutes'
									)}
									<div class="grid grid-cols-12 border-b border-black">
										<div class="col-span-5 grid grid-cols-12">
											<div class="px-2 border-l border-black uppercase col-span-6 py-1">
												{shift[1]}
											</div>
											<div class="px-2 border-l border-black uppercase col-span-6 py-1">
												{moment(
													'2023-12-12 ' + convertToMilitaryTime(shift[2].split('-')[0])
												).format('hh:mm A') +
													' - ' +
													moment(
														'2023-12-12 ' + convertToMilitaryTime(shift[2].split('-')[1])
													).format('hh:mm A')}
											</div>
										</div>
										<div class="col-span-7 grid grid-cols-12">
											<div class="col-span-7 grid grid-cols-3">
												<div class="px-2 border-l border-black uppercase py-1">&nbsp;</div>
												<div
													class:bg-gray-300={diff < 0 ? diff + 1440 <= 390 : diff <= 390}
													class="px-2 border-l border-black uppercase py-1 text-center"
												>
													{diff > 0 ? (diff <= 390 ? 'N/A' : '') : diff + 1440 <= 390 ? 'N/A' : ''}
												</div>
												<div
													class:bg-gray-300={diff < 0 ? diff + 1440 < 360 : diff < 360}
													class="px-2 border-l border-black uppercase py-1 text-center"
												>
													{diff > 0 ? (diff < 360 ? 'N/A' : '') : diff + 1440 < 360 ? 'N/A' : ''}
												</div>
											</div>
											<div class="col-span-5">
												<div class="px-2 border-l border-black border-r uppercase py-1">&nbsp;</div>
											</div>
										</div>
									</div>
								{/each}
							{/if}
						{/each}
					</div>
				{/each}
			</div>
		</div>
	{:else}
		<div class="mb-2 mt-4 flex justify-center">
			<img
				width="90"
				height="90"
				src="https://corporate.homedepot.com/sites/default/files/image_gallery/THD_logo.jpg"
				alt=""
			/>
		</div>
		<div class="text-[20pt] mb-3 font-bold text-orange-500">Daily Warboard Report Generator</div>
		<div class="flex justify-center">
			<div class="w-2/5 bg-orange-200 py-3">
				<input type="file" id="excel-file" accept=".xlsx, .xls" on:change={startLoad} />
			</div>
		</div>

		<div class="text-xl text-orange-400 mt-3">
			Upload .xlsx Dimensions Store Coverage By Day above to continue...
		</div>
	{/if}
</div>

<style>
	.hamster {
		display: inline-block;
		animation: bounce 1s ease-in-out infinite;
	}

	@keyframes bounce {
		0%,
		100% {
			transform: translateY(0);
		}
		50% {
			transform: translateY(-20px);
		}
	}
	@media print {
		.page-break {
			page-break-after: always;
		}
	}
	@media screen {
		.print-only {
			display: none;
		}
	}
	@media print {
		.screen-only {
			display: none;
		}
	}
</style>

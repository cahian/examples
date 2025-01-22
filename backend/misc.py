import subprocess, time
from argparse import ArgumentTypeError
from typing import (
	Any,
	Callable,
	Iterable,
	Iterator,
	List,
	Tuple,
	Union,
)

import psutil


def execute_command(command: str, raise_on_return_code: bool = True) -> None:
	"""
	**Execute a shell command as a subprocess and waits for it to finish**

	**Args:**
		* command (str):
			The shell command to execute.
		* raise_on_return_code (bool, optional):
			A flag indicating whether to raise an exception if the command
			returns a non-zero code. Defaults to True.

	**Returns:**
		* int:
			The return code of the command.

	**Exceptions**:
		* RuntimeError:
			If the command returns a non-zero code and raise_on_return_code is True.
	"""
	process = subprocess.Popen(
		command,
		shell=True,
		stdout=subprocess.PIPE,
		stderr=subprocess.PIPE,
		universal_newlines=True,
	)

	while True:
		stdout = process.stdout.readline()
		if stdout == '' and process.poll() is not None:
			break
		if stdout:
			print(stdout.strip())

	stdout, stderr = process.communicate()
	print(stdout)

	returncode = process.returncode
	if raise_on_return_code and returncode:
		if stderr:
			print('*************************************************')
			print('************ Subprocess error output ************')
			print(stderr)
			print('*************************************************')
			print('*************************************************')
		raise RuntimeError(f'Command "{command}" returned code {returncode}')
	return returncode


def monitor_memory(
	process_id: int,
	memory_limit: int,
	on_memory_exceeded: Callable,
	interval: float = 1.0,
) -> None:
	"""
	**Monitor the memory usage of a given process and calls a callback function
	if a limit is exceeded**

	**Args:**
		* process_id (int):
			The ID of the process to monitor.
		* memory_limit (int):
			The maximum allowed memory usage in bytes.
		* on_memory_exceeded (Callable):
			The function to call when the memory limit is exceeded. This function
			should accept the psutil.process as its only argument.
		* interval (float, optional):
			Time interval in seconds between memory checks. Defaults to 1.0.
	"""
	process = psutil.Process(process_id)
	while process.is_running():
		mem_info = process.memory_info().rss
		if mem_info > memory_limit:
			on_memory_exceeded(process)
			break
		time.sleep(interval)


def take_n(iterable: Iterable, length: int) -> Tuple[List, bool]:
	"""
	**Take the first n elements from an iterable**

	**Args:**
		* iterable (Iterable):
			The iterable to take elements from.
		* length (int):
			The number of elements to take.

	**Returns:**
		* Tuple[List, bool]:
			A tuple containing the first n elements from the iterable and a flag
			indicating whether the length of the iterable was greater than n.
	"""
	result = []
	for item in iterable:
		result.append(item)
		if len(result) == length:
			return result, True
	return result, False


def separate_equal_chunks(
	item_list: Iterable[Any], chunk_size: int, return_indexes: bool = False
) -> Iterator[Union[List[Any], Tuple[List[Any], int, int]]]:
	"""
	**Separate a list of items into equally sized chunks**

	**Args:**
		* item_list (Iterable):
			The list of items to be separated.
		* chunk_size (int):
			The desired size of each chunk.
		* return_indexes (bool, optional):
			A flag indicating whether to return the starting and ending indexes
			of each chunk. Defaults to False.

	**Yields**:
		* list or tuple:
			A list containing the items in the current chunk, or a tuple containing
			the chunk, its starting index, and its ending index if return_indexes
			is True.
	"""
	chunk = []
	chunk_start = 0
	for index, item in enumerate(item_list):
		chunk.append(item)
		if len(chunk) == chunk_size:
			if return_indexes:
				yield chunk, chunk_start, index
			else:
				yield chunk
			chunk = []
			chunk_start = index + 1
	# Yield the last remaining chunk (if any)
	if chunk:
		if return_indexes:
			yield chunk, chunk_start, chunk_start + len(chunk) - 1
		else:
			yield chunk

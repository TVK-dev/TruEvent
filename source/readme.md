## Introduction

The amount of data obtained by fluorescence microscopy of nerve cell cultures is very large both in terms of the number of cells studied and the time of observation, so the problem of algorithmic image processing to obtain some numerical metrics is relevant. For this, machine learning algorithms can be applied that classify objects and events. In order to train a machine learning algorithm to correctly classify objects and events, a person must first solve this task for it. This project is designed to facilitate the labeling of source data when solving machine learning problems.

## Project structure

Project consists of the following directories:

  1. `data` - directory with a small amount of test samples.
  2. `source` - source code.

## Installation / Getting Started

In VS6.0, open the TruEvent.vbp file

## Development

To start further development of the project, in the Project-TruEvent project tree, select the commform module containing the main functions

## Deploy/publish

The project contains a compiled exe file that does not require additional ocx elements, there should not be any difficulties.

## Functions

The working window of the calcium event editing program includes the following elements:

- area of graphs, which displays the intensity line, average and base lines, suggested and confirmed events
- list of files in which cell selection is made
- event parameters field, in which the parameters of the selected event are edited
- measurement field

The measurement field displays the measurement results:

- height from the top point of the peak to the corresponding baseline point below it for the selected event (absolute and relative values);
- height between the selected points of the intensity line (absolute and relative values);
- duration of the selected event in seconds.

By clicking on the graph area with the left mouse button, you can add a new event if it was skipped by the preliminary event detection algorithm.

By pressing the right mouse button on the graph area, the points between which the measurement is carried out are indicated. After the first click, the measurement field displays the inscription specify the 2nd point. Then, in the graph area, by pressing the right mouse button, you need to specify the position of the second point.

## Configuration

Data is prepared using a module implemented in Matlab

## Contribute

If you'd like to contribute, please fork the repository and use the feature branch. Pull requests are welcome

## Links

- Project homepage: https://github.com/TVK-dev/TruEvent
- Repository: https://github.com/TVK-dev/TruEvent
- Related projects:
   - Python library: https://github.com/TVK-dev/Intensity
   - Astrocyte Laboratory repository: [https://github.com/UNN-VMK-Software/astro-analysis](https://github.com/UNN-VMK-Software/astro-analysis)

## Licensing

The code in this project is licensed under the Attribution-NonCommercial 2.0 Generic license



## Введение

Одним из методов исследования высшей нервной деятельности является флюоресцентная микроскопия. Это связано с тем, что самые распространённые глиальные клетки – астроциты хоть и являются электрически невозбудимыми, но также участвуют в процессах, связанных с передачей сигналов, генерируя кальциевые сигналы. А кальциевые сигналы можно легко визуализировать при помощи флуоресцентных маркеров.

Объём данных, полученных при флюоресцентной микроскопии культур нервных клеток очень велик как по количеству изучаемых клеток, так и по времени наблюдения, поэтому актуальна задача алгоритмической обработки изображений для получения некоторых численных метрик. Для этого могут быть применены алгоритмы машинного обучения, осуществляющие классификацию объектов и событий.
Чтобы алгоритм машинного обучения обучить верно классифицировать объекты и события, за него это задачу сначала должен решить человек. Данный проект предназначен для облегчения разметки исходных данных при решении задач машинного обучения.

## Установка / Начало работы

В VS6.0 открыть файл TruEvent.vbp

## Разработка

Чтобы приступить к дальнейшей разработке проекта в дереве проекта Project-TruEvent выбрать модуль commform, содержащий основные функции

## Развертывание / публикация

Проект содержит откомпилированный exe файл, не требующий дополнительных элементов ocx, сложностей возникнуть не должно.

## Функции

Рабочее окно программы редактирования кальциевых событий включает в себя следующие элементы:

- область графиков, в которой отображаются линия интенсивности, средняя и базовая линии, предложенные и подтверждённые события
- список файлов, в котором производится выбор клетки
- поле параметров события, в котором редактируются параметры выбранного события
- поле измерений

В поле измерений отображаются результаты измерений:

- высота от верхней точки пика до соответствующей находящейся под ней точки базовой линии для выбранного события (абсолютное и относительное значения);
- высота между выбранными точками линии интенсивности (абсолютное и относительное значения);
- длительность выбранного события в секундах.

![](RackMultipart20210630-4-11tjysf_html_f0ebec70646eb330.png) ![](RackMultipart20210630-4-11tjysf_html_8a9953a29f4792c1.gif)

Нажатием по области графика левой клавиши мыши можно добавить новое событие если оно было пропущено алгоритмом предварительного детектирования событий.

Нажатием по области графика правой клавиши мыши производится указание точек, между которыми проводится измерение. После первого нажатия в поле измерений отображается надпись укажите 2 точку. Тогда в области графика нажатием правой клавиши мыши нужно указать положение второй точки.

## Конфигурация

Данные подготавдиваются при помощи модуля реализованного в Matlab

## Содействие

Если вы хотите внести свой вклад, пожалуйста, создайте ответвление репозитория и используйте функциональную ветку. Запросы на извлечение приветствуются

## Ссылки

- Домашняя страница проекта: https://github.com/TVK-dev/TruEvent
- Репозиторий: https://github.com/TVK-dev/TruEvent
- Связанные проекты:
  - Библиотека на python: https://github.com/TVK-dev/Intensity
  - Репозиторий «Astrocyte Laboratory»: [https://github.com/UNN-VMK-Software/astro-analysis](https://github.com/UNN-VMK-Software/astro-analysis)

## Лицензирование

Код в этом проекте находится под лицензией Attribution-NonCommercial 2.0 Generic

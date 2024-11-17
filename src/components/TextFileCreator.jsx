import React, { useState } from 'react';
import { saveAs } from 'file-saver';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import Svg from '../style/bg.svg';

const DocxEditor = () => {
    // Получаем текущий год, месяц и день
    const currentDate = new Date();
    const currentYear = currentDate.getFullYear();
    const currentMonth = currentDate.getMonth() + 1; // Месяцы от 0 до 11
    const currentDay = currentDate.getDate();
    const currentDayString = `${currentYear} ${currentMonth < 10 ? '0' + currentMonth : currentMonth} ${currentDay < 10 ? '0' + currentDay : currentDay}`;

    // Состояние для хранения всех данных (человек и спонсор)
    const [formData, setFormData] = useState({
        name: '',
        surname: '',
        fullNameKorean: '',
        birthYear: '',
        birthMonth: '',
        birthDay: '',
        passportNumber: '',
        phoneNumber: '',
        email: '',
        address: '',
        school: '',
        studyFromYear: '',
        studyFromMonth: '',
        studyToYear: '',
        studyToMonth: '',
        sponsorSurname: '',
        sponsorName: '',
        aboutYourSelf: '',
        WhoYourSponsor: '',
        SponsorOccupation: '',
        sponsorPhoneNumber: '',
        sponsorAddress: ''
    });

    // Состояние для хранения выбранного .docx файла
    const [docxFile, setDocxFile] = useState(null);

    // Функция для обработки изменения значений в инпуте
    const handleInputChange = (e) => {
        const { name, value } = e.target;
        setFormData({
            ...formData,
            [name]: value
        });
    };

    // Обработчик загрузки .docx файла
    const handleFileUpload = (event) => {
        const file = event.target.files[0];
        if (file && file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
            setDocxFile(file);
        } else {
            alert('Пожалуйста, загрузите файл формата .docx');
        }
    };

    // Функция для вычисления periodOfStudy в месяцах
    const calculatePeriodOfStudy = () => {
        const { studyFromYear, studyFromMonth, studyToYear, studyToMonth } = formData;

        if (studyFromYear && studyFromMonth && studyToYear && studyToMonth) {
            // Преобразуем года и месяца в Date объекты для вычисления разницы
            const startDate = new Date(studyFromYear, studyFromMonth - 1); // Месяцы от 0 до 11
            const endDate = new Date(studyToYear, studyToMonth - 1);

            // Получаем разницу в месяцах
            const yearsDiff = endDate.getFullYear() - startDate.getFullYear();
            const monthsDiff = endDate.getMonth() - startDate.getMonth();

            // Рассчитываем полное количество месяцев
            const totalMonths = yearsDiff * 12 + monthsDiff;

            return totalMonths;
        }
        return '';
    };

    // Обработчик замены текста в .docx файле
    const handleReplaceTextInDocx = async () => {
        // Проверка на загрузку файла и заполнение всех полей формы
        if (!docxFile) {
            alert('Пожалуйста, загрузите .docx файл');
            return;
        }

        if (Object.values(formData).includes('')) {
            alert('Пожалуйста, заполните все поля формы');
            return;
        }

        // Вычисляем periodOfStudy в месяцах
        const periodOfStudy = calculatePeriodOfStudy();

        // Создаем объект данных для отправки, добавляя вычисленные поля
        const updatedFormData = {
            ...formData,
            periodOfStudy: periodOfStudy, // Добавляем вычисленный период обучения
            currentYear: currentYear.toString(), // Добавляем текущий год
            currentDay: currentDayString // Добавляем текущий день
        };

        try {
            const zip = await loadDocxFile(docxFile);

            // Заменяем текст в шаблоне
            const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

            // Передаем данные для замены в шаблон
            doc.setData(updatedFormData);

            // Рендерим новый .docx
            doc.render();

            // Генерация нового .docx файла
            const newDocx = doc.getZip().generate({ type: 'blob' });

            // Скачиваем новый файл
            saveAs(newDocx, 'updated_document.docx');
        } catch (error) {
            console.error('Ошибка при создании документа:', error);
            alert('Произошла ошибка при создании документа.');
        }
    };

    // Функция для загрузки .docx файла
    const loadDocxFile = (file) => {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = function (e) {
                const data = e.target.result;
                const zip = new PizZip(data); // Используем PizZip для работы с архивом
                resolve(zip);
            };
            reader.onerror = reject;
            reader.readAsArrayBuffer(file);
        });
    };

    // Перевод меток на русский
    const translateLabel = (key) => {
        const labels = {
            name: 'Имя (английский)',
            surname: 'Фамилия (английский)',
            fullNameKorean: 'Полное имя (корейский)',
            birthYear: 'Год рождения',
            birthMonth: 'Месяц рождения',
            birthDay: 'День рождения',
            passportNumber: 'Номер паспорта',
            phoneNumber: 'Номер телефона',
            email: 'Электронная почта',
            address: 'Адрес',
            school: 'Школа',
            studyFromYear: 'Год начала учебы',
            studyFromMonth: 'Месяц начала учебы',
            studyToYear: 'Год окончания учебы',
            studyToMonth: 'Месяц окончания учебы',
            sponsorSurname: 'Фамилия спонсора',
            sponsorName: 'Имя спонсора',
            aboutYourSelf: 'О себе',
            WhoYourSponsor: 'Кто ваш спонсор?',
            SponsorOccupation: 'Род деятельности спонсора',
            sponsorPhoneNumber: 'Телефон спонсора',
            sponsorAddress: 'Адрес спонсора'
        };
        return labels[key] || key; // Если нет перевода, возвращаем исходный ключ
    };

    // Массив всех полей
    const inputFields = [
        { name: 'name', label: 'name' },
        { name: 'surname', label: 'surname' },
        { name: 'fullNameKorean', label: 'fullNameKorean' },
        { name: 'birthYear', label: 'birthYear' },
        { name: 'birthMonth', label: 'birthMonth' },
        { name: 'birthDay', label: 'birthDay' },
        { name: 'passportNumber', label: 'passportNumber' },
        { name: 'phoneNumber', label: 'phoneNumber' },
        { name: 'email', label: 'email' },
        { name: 'address', label: 'address' },
        { name: 'school', label: 'school' },
        { name: 'studyFromYear', label: 'studyFromYear' },
        { name: 'studyFromMonth', label: 'studyFromMonth' },
        { name: 'studyToYear', label: 'studyToYear' },
        { name: 'studyToMonth', label: 'studyToMonth' },
        { name: 'sponsorSurname', label: 'sponsorSurname' },
        { name: 'sponsorName', label: 'sponsorName' },
        { name: 'aboutYourSelf', label: 'aboutYourSelf' },
        { name: 'WhoYourSponsor', label: 'WhoYourSponsor' },
        { name: 'SponsorOccupation', label: 'SponsorOccupation' },
        { name: 'sponsorPhoneNumber', label: 'sponsorPhoneNumber' },
        { name: 'sponsorAddress', label: 'sponsorAddress' }
    ];

    return (
        <div className='main'>
            <div className='svg'><img src={Svg} alt="" /></div>

            {/* Инпуты для ввода данных */}
            <div className='mmm'>
                <div className='main_items'>
                    {inputFields.map((field) => (
                        <div className='item' key={field.name}>
                            <label className='rus'>{translateLabel(field.label)}</label>
                            <input
                                type="text"
                                name={field.name}
                                value={formData[field.name]}
                                onChange={handleInputChange}
                                placeholder={`Введите ${translateLabel(field.label)}`}
                            />
                        </div>
                    ))}
                </div>
            </div>

            {/* Загрузка .docx файла */}
            <input
                hidden
                type="file"
                accept=".docx"
                id="handleFileUpload"
                onChange={handleFileUpload}
            />

            {/* Кнопка для открытия выбора файла */}
            <button className='btn' onClick={() => document.getElementById('handleFileUpload').click()}>Выберите файл KIU.docx</button>

            {/* Кнопка для замены текста в .docx */}
            <button className='btn' style={{ bottom: '80px' }} onClick={handleReplaceTextInDocx}>Заменить данные в .docx</button>
        </div>
    );
};

export default DocxEditor;

import React, { useState } from 'react';
import { saveAs } from 'file-saver';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import Svg from '../style/bg.svg';

const DocxEditor = () => {
    // Получаем текущий год динамически
    const currentYear = new Date().getFullYear();

    // Состояния для хранения значений из инпутов
    const [docxFile, setDocxFile] = useState(null);

    // Данные для динамических инпутов с начальными значениями
    const [formData, setFormData] = useState({
        currentYear: currentYear.toString(), // Текущий год
        name: 'ABDUMALIK',
        surname: 'DZHOROBAEV',
        fullNameKorean: '조로바예프 압두말리크',
        birthYear: '2007',
        birthMonth: '04',
        birthDay: '09',
        passportNumber: 'PE1292077',
        phoneNumber: '+996551014109',
        email: 'abdumalikdzhorobaev@gmail.com',
        address: 'Kyrgyzstan, Osh region',
        school: 'Ubileynaya',
        studyFromYear: '2012',
        studyToYear: '2023',
        fatherSurname: 'ABDURAZAKOV',
        fatherName: 'SHAVKAT',
        aboutYourSelf: 'i am human dwqdqwd qwd qw d qwd wq dqwdqwdqwd qw dqw d qw dq wd qw d qwdqwjdnjqwndnqjdwq dq wdjqwndjqwdnjqw dqw d qwd qw dqw  dqwndjqwndnqjwdnjqwd qw dq wd qw dw '
    });

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

        try {
            const zip = await loadDocxFile(docxFile);

            // Заменяем текст в шаблоне
            const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

            // Передаем данные для замены в шаблон
            doc.setData(formData);

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

    // Функция для определения типа инпута в зависимости от ключа
    const getInputType = (key) => {
        switch (key) {
            case 'email':
                return 'email';
            case 'phoneNumber':
                return 'tel';
            case 'birthYear':
            case 'birthMonth':
            case 'birthDay':
            case 'studyFromYear':
            case 'studyToYear':
            case 'currentYear':
                return 'number';
            default:
                return 'text';
        }
    };

    // Перевод меток на русский
    const translateLabel = (key) => {
        const labels = {
            currentYear: 'Текущий год',
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
            studyToYear: 'Год окончания учебы',
            fatherSurname: 'Фамилия отца',
            fatherName: 'Имя отца',
            aboutYourSelf: 'О себе'
        };
        return labels[key] || key; // Если нет перевода, возвращаем исходный ключ
    };

    return (
        <div className='main'>
            <div className='svg'><img src={Svg} alt="" /></div>

            {/* Инпуты для ввода данных */}
            <div className='mmm'>
                <div className='main_items'>
                    {Object.keys(formData).map((key) => (
                        <div className='item' key={key}>
                            <label className='rus'>{translateLabel(key)}</label> {/* Переводим метки */}
                            <input
                                type={getInputType(key)} // Динамически устанавливаем тип
                                name={key} // Имя инпута, которое будет соответствовать ключу в состоянии
                                onChange={handleInputChange} // Обработчик изменения значения
                                placeholder={`Введите ${translateLabel(key)}`} // Переводим placeholder
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
            <button className='btn' onClick={() => document.getElementById('handleFileUpload').click()}>Выберите файл жен KIU.docx</button>
            <button className='btn' style={{
                bottom: '80px'
            }} onClick={() => document.getElementById('handleFileUpload').click()}>Выберите файл муж KIU.docx</button>

            {/* Кнопка для замены текста в .docx */}
            <button style={{
                bottom: '170px'
            }} className='btn' onClick={handleReplaceTextInDocx}>Заменить данные в .docx</button>
        </div>
    );
};

export default DocxEditor;

<template>
  <div class="container mx-auto py-8 max-w-sm">
    <h1 class="text-3xl font-bold mb-2 text-center">Generate Surat Perintah Kerja Lembur</h1>
    <form @submit.prevent="generateDocument" class="bg-white rounded-lg shadow-lg p-6">
      <div class="mb-2">
        <label for="nama" class="text-sm font-medium">Nama:</label>
        <input id="nama" type="text" v-model="nama" required
          class="w-full h-8 border-gray-300 mt-1 px-3 py-2 rounded-md focus:outline-none focus:ring-blue-500 focus:border-blue-500">
      </div>

      <div class="mb-2">
        <label for="nik" class="text-sm font-medium">NIK:</label>
        <input id="nik" type="text" v-model="nik" required
          class="w-full h-8 border-gray-300 mt-1 px-3 py-2 rounded-md focus:outline-none focus:ring-blue-500 focus:border-blue-500">
      </div>
      <div class="mb-2">
        <label for="keperluan" class="text-sm font-medium">Keperluan:</label>
        <input id="keperluan" type="text" v-model="keperluan" required
          class="w-full h-8 border-gray-300 mt-1 px-3 py-2 rounded-md focus:outline-none focus:ring-blue-500 focus:border-blue-500">
      </div>

      <div class="mb-2">
        <label for="tanggal_lembur" class="text-sm font-medium">Tanggal Lembur:</label>
        <input id="tanggal_lembur" type="date" v-model="tanggalLembur" required
          class="w-full h-8 border-gray-300 mt-1 px-3 py-2 rounded-md focus:outline-none focus:ring-blue-500 focus:border-blue-500">
      </div>
      <div class="mb-2">
        <label for="tanggal_acc" class="text-sm font-medium">Tanggal Acc:</label>
        <input id="tanggal_acc" type="date" v-model="tanggal_acc" required
          class="w-full h-8 border-gray-300 mt-1 px-3 py-2 rounded-md focus:outline-none focus:ring-blue-500 focus:border-blue-500">
      </div>

      <div class="mb-2">
        <label for="time_start" class="text-sm font-medium">Waktu Mulai:</label>
        <input id="time_start" type="time" v-model="timeStart" required
          class="w-full h-8 border-gray-300 mt-1 px-3 py-2 rounded-md focus:outline-none focus:ring-blue-500 focus:border-blue-500">
      </div>

      <div class="mb-2">
        <label for="time_end" class="text-sm font-medium">Waktu Selesai:</label>
        <input id="time_end" type="time" v-model="timeEnd" required
          class="w-full h-8 border-gray-300 mt-1 px-3 py-2 rounded-md focus:outline-none focus:ring-blue-500 focus:border-blue-500">
      </div>

      <button type="submit"
        class="w-full h-8 bg-blue-500 text-white font-semibold py-2 px-4 rounded hover:bg-blue-600 focus:outline-none focus:bg-blue-600">
        Generate
      </button>
    </form>
  </div>
</template>


<script setup>
import { ref } from 'vue';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import { format, differenceInHours, differenceInMinutes } from 'date-fns';
import { id } from 'date-fns/locale';
import { saveAs } from 'file-saver';

const nama = ref('');
const nik = ref('');
const keperluan = ref('');
const tanggal_acc = ref('');
const tanggalLembur = ref('');
const timeStart = ref('');
const timeEnd = ref('');

const generateDocument = () => {
  // Path to the template file in the public folder
  const templatePath = '/spkl.docx';

  // Fetch the template file
  fetch(templatePath)
    .then(response => response.arrayBuffer())
    .then(templateBuffer => {
      const zip = new PizZip(templateBuffer);
      const doc = new Docxtemplater();
      doc.loadZip(zip);

      // Format the date as dd-mm-yyyy
      const formattedDate = format(new Date(tanggalLembur.value), "EEEE, dd MMMM yyyy", { locale: id });
      const formattedAccDate = format(new Date(tanggal_acc.value), "EEEE, dd MMMM yyyy", { locale: id });

      // Calculate the distance in hours
      const distanceInMinutes = differenceInMinutes(new Date(`2000-01-01 ${timeEnd.value}`), new Date(`2000-01-01 ${timeStart.value}`));
      const distanceInHours = distanceInMinutes / 60;

      // Set the data for the document
      const data = {
        nama: nama.value,
        nik: nik.value,
        tanggal_lembur: formattedDate,
        waktu_mulai: timeStart.value,
        waktu_selesai: timeEnd.value,
        keperluan: keperluan.value,
        tanggal_acc : formattedAccDate,
        durasi: distanceInHours,
        header: 'Generate Surat Perintah Kerja Lembur'
      };
      doc.setData(data);

      // Perform the document generation
      try {
        doc.render();
      } catch (error) {
        const e = {
          message: error.message,
          name: error.name,
          stack: error.stack,
          properties: error.properties
        };
        console.error('Error rendering document:', e);
        return;
      }

      const generatedDocument = doc.getZip().generate({ type: 'blob' });
      const outputFileName = `SPKL_${nama.value}_${formattedDate}.docx`;

      // Use FileSaver.js to save the generated document
      saveAs(generatedDocument, outputFileName);
    })
    .catch(error => {
      console.error('Error fetching template file:', error);
    });
};

</script>


<style>
body {
  background-color: #f3f4f6;
}
</style>

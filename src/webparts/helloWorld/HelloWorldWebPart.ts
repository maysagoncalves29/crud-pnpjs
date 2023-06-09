import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'HelloWorldWebPartStrings';
import styles from './styles.scss';

import pnp, { sp } from 'sp-pnp-js';
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";


export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
  private _selectedItemId: any;

  constructor() {
    super();
    const params = new URLSearchParams(location.search);
    this._selectedItemId = params.get("itemId");
    this.readAluno();

    
    
  }


  public render(): void {
      this.domElement.innerHTML = `
    <div class="${styles}"></div>
    <div class="formulario">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.3/css/all.min.css" />
  <form>

    <div class="form-row">
      <label for="name">Nome:</label>
      <input type="text" id="name" placeholder="Digite o nome do Aluno.">
    </div>
    <div class="form-row">
      <label for="escolas">Escolas:</label>
      <select id="escolas" placeholder="Selecione a escola do Aluno."></select>
    </div>
    <div class="form-row">
      <label for="email">Email:</label>
      <input type="text" id="email" placeholder="Digite o e-mail do responsável ou do Aluno">
    </div>
    <div class="form-row">
      <label for="aprovado">Aprovado? Se o aluno foi aprovado, selecionar opção abaixo. </label>
      <input type="checkbox" id="aprovado">
    </div>
    <div class="form-row">
      <label for="alunoSala">Sala:</label>
      <div class="custom-select">
      <select id="alunoSala" multiple >
        <option value="Fundamental I">Fundamental I</option>
        <option value="Fundamental II">Fundamental II</option>
        <option value="Ensino Médio">Ensino Médio</option>
      </select>
      <div class="select-dropdown"></div>
    </div>
    </div>
    <div class="form-row">
      <label for="alunoCidade">Cidade:</label>
      <input type="text" id="alunoCidade" placeholder="Digite a cidade do Aluno.">
    </div>
    <div class="form-row">
    <button type="submit" id="btnUpdate" ${this._selectedItemId ? '' : 'style="display:none;"'}><i class="fas fa-pencil-alt"></i> Atualizar Aluno</button>
    <button type="submit" id="btnDelete" ${this._selectedItemId ? '' : 'style="display:none;"'}><i class="fas fa-trash"></i> Remover Aluno</button>
    <button type="submit" id="btnInsert" ${this._selectedItemId ? 'style="display:none;"' : ''}><i class="fas fa-plus"></i> Cadastrar Aluno</button>
    
    </div>
  </form>
</div>

    `;
    
    sp.web.lists.getByTitle("Escola").items.select("Title", "Id").get().then((items: any[]) => {
      const dropdown = document.getElementById("escolas") as HTMLSelectElement;
      for (const item of items) {
        const option = document.createElement("option");
        option.value = item.Id;
        option.text = item.Title;
        dropdown.add(option);
      }
    }).catch((error: any) => {
      console.log(error);
    });
    
    this.bindEvent();
    this.readAluno();
  }

 

  
  private bindEvent(): void {
    this.domElement.querySelector('#btnInsert').addEventListener('click', (event) => {
      event.preventDefault();
      this.insertAluno();
    });
    this.domElement.querySelector('#btnUpdate').addEventListener('click', (event) => {
      event.preventDefault();
      this.updateAluno();
    });
    this.domElement.querySelector('#btnDelete').addEventListener('click', (event) => {
      event.preventDefault();
      this.deleteAluno();
    });
  }
 
  

  private insertAluno() : void {
    const Name : string = (document.getElementById("name") as HTMLInputElement).value
    const Email : string = (document.getElementById("email") as HTMLInputElement).value;
    const Aprovado : boolean = (document.getElementById("aprovado") as HTMLInputElement).checked;
    const AlunoCidade : string = (document.getElementById("alunoCidade") as HTMLInputElement).value;
    const Escolas : string = (document.getElementById("escolas") as HTMLSelectElement).value;  
    let AlunoSala = (document.getElementById("alunoSala") as HTMLSelectElement).selectedOptions;
    let OpcoesSala = [];
    for(let i = 0; i < AlunoSala.length; i++){
      OpcoesSala.push(AlunoSala[i].value);
    }
    console.log(OpcoesSala);
    sp.web.lists.getByTitle("Students").items.add({
      Title: AlunoCidade,
      Name: Name,
      Email: Email,
      Aprovado: Aprovado,
      AlunoSala: { results: OpcoesSala },
      EscolasId:  parseInt(Escolas),
      //AlunoSalaChoices: { results: choices },
      //Responsavel: parseInt(Responsavel)     
    }).then((_response: unknown) => { 
      alert('Aluno cadastrado com Sucesso!');
    }).catch((error: unknown) => {
      console.log(error);
    });
  }
  private updateAluno() : void {
    const Name : string = (document.getElementById("name") as HTMLInputElement).value
    const Email : string = (document.getElementById("email") as HTMLInputElement).value;
    const Aprovado : boolean = (document.getElementById("aprovado") as HTMLInputElement).checked;
    const AlunoCidade : string = (document.getElementById("alunoCidade") as HTMLInputElement).value;
    const Escolas : string = (document.getElementById("escolas") as HTMLSelectElement).value;
    let AlunoSala = (document.getElementById("alunoSala") as HTMLSelectElement).selectedOptions;
    let OpcoesSala = [];
    for(let i = 0; i < AlunoSala.length; i++){
      OpcoesSala.push(AlunoSala[i].value);
    }
    console.log(OpcoesSala);

    sp.web.lists.getByTitle("Students").items.getById(this._selectedItemId).update({

      //sp.web.lists.getByTitle("Students").fields.addMultiChoice("AlunoSala", {Choices: choices}), 
      Title: AlunoCidade,
      Name: Name,
      Email: Email,
      Aprovado: Aprovado,
      AlunoSala: { results: OpcoesSala  },
      EscolasId: parseInt(Escolas),


    }).then((_response: unknown) => { 
      alert('Aluno atualizado com sucesso!');
    }).catch((error: unknown) => { 
      console.log(error);
      this.displayMessage("Nao foi!")
    });

  }
  private deleteAluno() : void  {

    
    sp.web.lists.getByTitle("Students").items.getById(this._selectedItemId).delete().then(() => {
      console.log("Aluno excluído com sucesso");
      this.displayMessage("Aluno excluído com sucesso");
  }).catch((error) => {
      console.log(`Ocorreu um erro ao excluir o item: ${error}`);
      this.displayMessage("Não foi!");
  });

}
  private readAluno() : void {
    if(this._selectedItemId == null) {
      sp.web.lists.getByTitle("Escola").items.get().then((response) => {
        const escolasSelect = document.getElementById("escolas") as HTMLSelectElement;
        escolasSelect.innerHTML = "";
        response.forEach((item : any) => {
          const option = document.createElement("option");
          option.value = item.Id.toString();
          option.text = item.Title;
          escolasSelect.add(option);
        });
      }).catch((error) => {
        console.log(error);
      });
    } else {
        sp.web.lists.getByTitle("Students").items.getById(this._selectedItemId).select("Aprovado", "Email", "Name", "AlunoSala", "Title", "Escolas/Title", "Escolas/Id").expand("Escolas").get().then((item) => {
          
        
      
          // eslint-disable-next-line no-void, @typescript-eslint/no-floating-promises
          sp.web.lists.getByTitle("Escola").items.get().then((response) => {
            const escolasSelect = document.getElementById("escolas") as HTMLSelectElement;
            escolasSelect.innerHTML = "";
            response.forEach((item : any) => {
              const option = document.createElement("option");
              option.value = item.Id.toString();
              option.text = item.Title;
              escolasSelect.add(option);
            });
      if (escolasSelect) {
        escolasSelect.value = item.Escolas.Id;
      }
          }).catch((error) => {
            console.log(error);
          });
      console.log("foi");
      let saida: string = "";
      for (let i: number = 0; i < item.length; i++) { 
        saida += `Name: ${item[i].Name}, Email: ${item[i].Email}, Aprovado: ${item[i].Aprovado}, Aluno Sala: ${item[i].AlunoSala}, Aluno Cidade: ${item[i].Title}, Escolas: ${item[i].Escolas}\n`; 
      }



      const nameElement = document.getElementById("name") as HTMLInputElement;
      if (nameElement) {
        nameElement.value = item.Name; 
      }
  
      const emailElement = document.getElementById("email") as HTMLInputElement;
      if (emailElement) {
        emailElement.value = item.Email;
      }
  
      const aprovadoElement = document.getElementById("aprovado") as HTMLInputElement;
      if (aprovadoElement) {
        aprovadoElement.checked = item.Aprovado;
      }
  
      const alunoSalaElement = document.getElementById("alunoSala") as HTMLSelectElement;
      if (alunoSalaElement) {
        //const values = ["Fundamental I", "Fundamental II", "Ensino Médio"];
        const values = item.AlunoSala
        for (let i = 0; i < alunoSalaElement.options.length; i++) {
          const option = alunoSalaElement.options[i];
          option.selected = values.indexOf(option.value) >= 0;   
      }
      //alunoSalaElement.value = item.AlunoSala;
      //console.log("oie")
       
      }
  
      const alunoCidadeElement = document.getElementById("alunoCidade") as HTMLInputElement;
      if (alunoCidadeElement) {
        alunoCidadeElement.value = item.Title;
      }
      
      
    console.log(item);
    document.getElementById("MsgStatus").innerText = saida;
    return saida;
    //this.renderTable(this._selectedItemId);
    }).catch((error) => { //testando
    console.log(error);
    });
  }
}

  
  



  displayMessage(message: string) {
    alert(message)
  }
  // eslint-disable-next-line @typescript-eslint/explicit-function-return-type
  
  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
  
      pnp.setup({
        spfxContext: this.context
      });
    })



}


  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}


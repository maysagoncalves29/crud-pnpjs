import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'HelloWorldWebPartStrings';

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
    <div>
  <div>
    <table border='5'>
      <tr>
        <td>Name</td>
        <td><input type='text' id='name'/></td>
      </tr>
        <td>Escolas</td>
        <td>
          <select id="escolas">
          </select>
        </td>
      </tr>
      <tr>
        <td>Aluno Email</td>
        <td><input type='text' id='email'/></td>
        
      </tr>
      <tr>
        <td>Aluno Aprovado</td>
        <td><input type='checkbox' id='aprovado'/></td>
      </tr>
      <tr>
        <td>Aluno Sala</td>
        <td>
          <select id='alunoSala' multiple>
            <option value='Fundamental I'>Fundamental I</option>
            <option value='Fundamental II'>Fundamental II</option>
            <option value='Ensino Médio'>Ensino Médio</option>
          </select>
        </td>
      </tr>
      <tr>
        <td>Aluno Cidade</td>
        <td><input type='text' id='alunoCidade'/></td>
      </tr>
      <tr>
        <td>
          <input type='submit' value='Insert' id='btnInsert'/>
          <input type='submit' value='Update' id='btnUpdate'/>
          <input type='submit' value='Delete' id='btnDelete'/>
        </td>
      </tr>
    </table>
  </div>
  <div id="MsgStatus"></div>
</div>

    `;
    this.bindEvent();
    this.readAluno();
    
  
   
  }

 

  
  private bindEvent(): void {
    this.domElement.querySelector('#btnInsert').addEventListener('click', () => { this.insertAluno(); });
    this.domElement.querySelector('#btnUpdate').addEventListener('click', () => { this.updateAluno(); });
    this.domElement.querySelector('#btnDelete').addEventListener('click', () => { this.deleteAluno(); });
  
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
      alert('Add: Success!');
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
      //AlunoSalaChoices: { results: choices },
     //ResponsavelId: parseInt(Responsavel)


    }).then((_response: unknown) => { 
      alert('Update: Success!');
    }).catch((error: unknown) => { 
      console.log(error);
      this.displayMessage("Nao foi!")
    });

  }
  private deleteAluno() : void  {

    
    sp.web.lists.getByTitle("Students").items.getById(this._selectedItemId).delete().then(() => {
      console.log("Item excluído com sucesso");
      this.displayMessage("Item deletado com sucesso!");
  }).catch((error) => {
      console.log(`Ocorreu um erro ao excluir o item: ${error}`);
      this.displayMessage("Não foi!");
  });

}
  private readAluno() : void {
    


    //sp.web.lists.getByTitle("Students").items.getById(this._selectedItemId).select( "Aprovado", "Email", "Name", "AlunoSala", "Title", "Escolas/Title", "Participativos").expand("Escolas").get().then((item) => {
      //sp.web.lists.getByTitle("Students").items.getById(this._selectedItemId).select("AlunoSala").get().then((item) => {
        sp.web.lists.getByTitle("Students").items.getById(this._selectedItemId).select("Aprovado", "Email", "Name", "AlunoSala", "Title", "Escolas/Title").expand("Escolas").get().then((item) => {
      
          // eslint-disable-next-line no-void, @typescript-eslint/no-floating-promises
          sp.web.lists.getByTitle("Escola").items.select("Title", "Id").getAll().then((escolas) => {
            let select = document.getElementById('escolas');
            escolas.forEach((escola) => {
              let option = document.createElement('option');
              option.text = escola.Title;
              option.value = escola.Id;
              select.appendChild(option)
            })
          })
      
      let saida: string = "";
      for (let i: number = 0; i < item.length; i++) { 
        saida += `Name: ${item[i].Name}, Email: ${item[i].Email}, Aprovado: ${item[i].Aprovado}, Aluno Sala: ${item[i].AlunoSala}, Aluno Cidade: ${item[i].Title}, Escolas: ${item[i].EscolasId.Title}\n`; 
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
        alunoSalaElement.value = item.AlunoSala;
      }
  
      const alunoCidadeElement = document.getElementById("alunoCidade") as HTMLInputElement;
      if (alunoCidadeElement) {
        alunoCidadeElement.value = item.Title;
      }
      const escolasElement = document.getElementById("escolas") as HTMLSelectElement;
      if (escolasElement) {
        escolasElement.value = item.EscolasId.Title;
      }
      
    console.log(item);
    document.getElementById("MsgStatus").innerText = saida;
    return saida;
    //this.renderTable(this._selectedItemId);
    }).catch((error) => { //testando
    console.log(error);
    });
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

